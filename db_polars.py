"""
pudbo-polars — Upload-driven analytics dashboard with a Polars backend.
================================================================================
A companion to db_fw12.py. Unlike db_fw12 (a single self-contained HTML where ALL
data manipulation happens in browser JS), this is a small FastAPI server: the
browser only renders, and **Polars does every transform** server-side —
filter / sort / paginate / summary / pivot / group-by / join.

Relationships between tables map onto Polars joins with cardinality validation:
  one-to-one  → validate="1:1"
  one-to-many → validate="1:m"
  many-to-one → validate="m:1"
  many-to-many→ validate="m:m"
Joins support multi-column key mappings and a "keep columns" selection.

Group-by supports multiple metric+aggregate pairs at once.

The chart builder is ported from db_fw12.py — the full feature set (15 chart
types, combo dual-axis, palettes, trendlines, per-chart filters, labels, etc.).
Charts are rendered client-side from rows the Polars backend streams via /api/raw.

RUN
---
  pip install fastapi "uvicorn[standard]" python-multipart polars
  python db_polars_app.py
  # then open http://127.0.0.1:8000

Polars is already required for the manipulation; FastAPI/uvicorn host the API.
"""
from __future__ import annotations

import ast
import io
import json
import math
import secrets
import shutil
from pathlib import Path
from typing import Any

import polars as pl

try:
    from fastapi import FastAPI, File, HTTPException, Request, UploadFile
    from fastapi.responses import HTMLResponse, Response
except ModuleNotFoundError as e:  # pragma: no cover - friendly hint
    raise SystemExit(
        "Missing web deps. Install them with:\n"
        '  pip install fastapi "uvicorn[standard]" python-multipart\n'
        f"(original error: {e})"
    )

app = FastAPI(title="pudbo-polars")

# In-memory table store. name -> DataFrame, plus light metadata for the UI.
TABLES: dict[str, pl.DataFrame] = {}
META: dict[str, dict] = {}
# Saved charts per table: name -> [{"id","title","cfg"}, ...]
CHARTS: dict[str, list[dict]] = {}

# ──────────────────────────────────────────────────────────────────────────────
# Helpers — every data operation below runs through Polars
# ──────────────────────────────────────────────────────────────────────────────
_NUM_PREFIXES = ("Int", "UInt", "Float", "Decimal")


def _is_num(dtype) -> bool:
    return str(dtype).startswith(_NUM_PREFIXES)


def _unique_name(base: str) -> str:
    base = (base or "table").strip() or "table"
    name, i = base, 2
    while name in TABLES:
        name, i = f"{base}_{i}", i + 1
    return name


def _sanitize(v: Any):
    """Make a Polars cell JSON-safe."""
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, str)):
        return v
    if isinstance(v, float):
        return None if (math.isnan(v) or math.isinf(v)) else v
    return str(v)  # dates, datetimes, durations, structs, lists…


def _schema(df: pl.DataFrame) -> list[dict]:
    return [
        {"name": c, "dtype": str(df.schema[c]), "numeric": _is_num(df.schema[c])}
        for c in df.columns
    ]


def _numeric_cols(df: pl.DataFrame) -> list[str]:
    return [c for c in df.columns if _is_num(df.schema[c])]


def _summary(df: pl.DataFrame) -> list:
    """Column-aligned totals: sum for numeric columns, None elsewhere."""
    ncols = _numeric_cols(df)
    smap: dict[str, Any] = {}
    if ncols and df.height:
        totals = df.select([pl.col(c).sum().alias(c) for c in ncols]).row(0)
        smap = dict(zip(ncols, totals))
    return [_sanitize(smap.get(c)) if c in smap else None for c in df.columns]


def _agg_expr(col: str, agg: str) -> pl.Expr:
    e = pl.col(col)
    table = {
        "sum": e.sum(), "mean": e.mean(), "avg": e.mean(),
        "min": e.min(), "max": e.max(), "count": e.count(),
        "median": e.median(), "first": e.first(), "last": e.last(),
        "std": e.std(), "n_unique": e.n_unique(),
    }
    return table.get(agg, e.sum())


def _apply_query(df: pl.DataFrame, q: dict) -> pl.DataFrame:
    """Apply filters → search → sort, all in Polars."""
    for f in (q.get("filters") or []):
        col, op, val = f.get("col"), f.get("op"), f.get("value")
        if col not in df.columns:
            continue
        is_num = _is_num(df.schema[col])
        try:
            if op == "in" and isinstance(val, (list, tuple)):
                wanted = [str(x) for x in val]
                df = df.filter(pl.col(col).cast(pl.Utf8).is_in(wanted))
            elif op == "range" and is_num and isinstance(val, dict):
                lo, hi = val.get("min"), val.get("max")
                if lo is not None and lo != "":
                    df = df.filter(pl.col(col) >= float(lo))
                if hi is not None and hi != "":
                    df = df.filter(pl.col(col) <= float(hi))
            elif op == "between" and is_num and isinstance(val, (list, tuple)) and len(val) == 2:
                lo, hi = float(val[0]), float(val[1])
                df = df.filter((pl.col(col) >= lo) & (pl.col(col) <= hi))
            elif op == "contains":
                df = df.filter(
                    pl.col(col).cast(pl.Utf8).str.to_lowercase().str.contains(str(val).lower(), literal=True)
                )
            elif op in ("=", "!=", ">", ">=", "<", "<="):
                if is_num:
                    v = float(val)
                    col_e = pl.col(col)
                else:
                    v = str(val)
                    col_e = pl.col(col).cast(pl.Utf8)
                cmp = {
                    "=": col_e == v, "!=": col_e != v,
                    ">": col_e > v, ">=": col_e >= v,
                    "<": col_e < v, "<=": col_e <= v,
                }[op]
                df = df.filter(cmp)
        except Exception:
            pass  # bad filter input → ignore rather than 500

    search = str(q.get("search") or "").strip().lower()
    if search:
        df = df.filter(
            pl.any_horizontal(
                [pl.col(c).cast(pl.Utf8).str.to_lowercase().str.contains(search, literal=True) for c in df.columns]
            )
        )

    sort = [s for s in (q.get("sort") or []) if s.get("col") in df.columns]
    if sort:
        df = df.sort(
            by=[s["col"] for s in sort],
            descending=[int(s.get("dir", 1)) < 0 for s in sort],
            nulls_last=True,
        )
    return df


def _payload(df: pl.DataFrame, page: int = 1, page_size: int = 50) -> dict:
    total = df.height
    pages = max(1, math.ceil(total / page_size)) if page_size else 1
    page = min(max(1, int(page)), pages)
    sl = df.slice((page - 1) * page_size, page_size)
    rows = [[_sanitize(v) for v in r] for r in sl.rows()]
    return {
        "cols": df.columns,
        "schema": _schema(df),
        "numeric_cols": _numeric_cols(df),
        "rows": rows,
        "summary": _summary(df),
        "total": total,
        "page": page,
        "pages": pages,
        "page_size": page_size,
    }


def _register(df: pl.DataFrame, base: str, kind: str, detail: str = "") -> str:
    name = _unique_name(base)
    TABLES[name] = df
    META[name] = {"kind": kind, "detail": detail}
    _persist()
    return name


# ──────────────────────────────────────────────────────────────────────────────
# Undo history — every in-place transform snapshots the prior DataFrame
# ──────────────────────────────────────────────────────────────────────────────
HISTORY: dict[str, list[dict]] = {}   # name -> [{"df", "label"}, ...]  (most-recent last)
RECIPE: dict[str, list[dict]] = {}    # name -> [{"op", "params", "label"}, ...]  (replayable)
_HISTORY_MAX = 25


def _commit(name: str, new_df: pl.DataFrame, label: str):
    """Replace a table in place, pushing the previous state onto its undo stack."""
    hist = HISTORY.setdefault(name, [])
    hist.append({"df": TABLES[name], "label": label})
    if len(hist) > _HISTORY_MAX:
        del hist[0]
    TABLES[name] = new_df
    META.setdefault(name, {})["detail"] = label
    _persist()


# ──────────────────────────────────────────────────────────────────────────────
# Disk persistence — tables + charts survive a server restart
# ──────────────────────────────────────────────────────────────────────────────
STORE = Path(__file__).parent / "pudbo_store"
PROJECTS = STORE / "projects"          # one isolated sub-folder per project
_loading = False                       # suppress re-persist while loading at startup
_active_pid = "default"                # folder name of the currently-open project
_active_name = "Default"               # its human-readable display name


def _safe_file(name: str) -> str:
    return secrets.token_hex(6) if not name else "".join(c if c.isalnum() else "_" for c in name)[:40]


def _proj_dir(pid: str | None = None) -> Path:
    return PROJECTS / (pid or _active_pid)


def _slug(name: str) -> str:
    s = "-".join(filter(None, "".join(c if c.isalnum() else "-" for c in (name or "").strip().lower()).split("-")))
    return s[:40] or secrets.token_hex(4)


def _unique_pid(name: str) -> str:
    base = _slug(name)
    existing = {p.name for p in PROJECTS.iterdir()} if PROJECTS.exists() else set()
    pid, i = base, 2
    while pid in existing:
        pid, i = f"{base}-{i}", i + 1
    return pid


def _write_active():
    try:
        STORE.mkdir(exist_ok=True)
        (STORE / "_active.txt").write_text(_active_pid, encoding="utf-8")
    except Exception:
        pass


def _persist():
    """Save the OPEN project's tables + charts + recipes into its own folder."""
    if _loading:
        return
    try:
        d = _proj_dir()
        d.mkdir(parents=True, exist_ok=True)
        manifest = {"name": _active_name, "tables": [], "charts": CHARTS, "recipe": RECIPE}
        files = set()
        for n in TABLES:
            fn = _safe_file(n) + ".parquet"
            # avoid filename collisions between similarly-named tables
            while fn in files:
                fn = secrets.token_hex(4) + ".parquet"
            files.add(fn)
            TABLES[n].write_parquet(d / fn)
            manifest["tables"].append({"name": n, "file": fn, "meta": META.get(n, {})})
        (d / "_session.json").write_text(json.dumps(manifest), encoding="utf-8")
        # drop parquet files no longer referenced
        for p in d.glob("*.parquet"):
            if p.name not in files:
                p.unlink(missing_ok=True)
        _write_active()
    except Exception as e:  # persistence is best-effort, never fatal
        print(f"[persist] warning: {e}")


def _load_project(pid: str):
    """Replace ALL in-memory state with the contents of project <pid>."""
    global _loading, _active_pid, _active_name
    TABLES.clear(); META.clear(); CHARTS.clear(); HISTORY.clear(); RECIPE.clear()
    _active_pid = pid
    _active_name = pid
    sess = _proj_dir(pid) / "_session.json"
    if not sess.exists():
        return
    try:
        _loading = True
        manifest = json.loads(sess.read_text(encoding="utf-8"))
        _active_name = manifest.get("name", pid)
        for t in manifest.get("tables", []):
            p = _proj_dir(pid) / t["file"]
            if p.exists():
                TABLES[t["name"]] = pl.read_parquet(p)
                META[t["name"]] = t.get("meta", {"kind": "upload", "detail": ""})
        CHARTS.update(manifest.get("charts", {}))
        RECIPE.update(manifest.get("recipe", {}))
        print(f"[load] project '{_active_name}' — {len(TABLES)} table(s)")
    except Exception as e:
        print(f"[load] warning: {e}")
    finally:
        _loading = False


def _list_projects() -> list[dict]:
    out = []
    if PROJECTS.exists():
        for d in sorted(PROJECTS.iterdir()):
            if not d.is_dir():
                continue
            name, ntables = d.name, 0
            sess = d / "_session.json"
            if sess.exists():
                try:
                    m = json.loads(sess.read_text(encoding="utf-8"))
                    name = m.get("name", d.name)
                    ntables = len(m.get("tables", []))
                except Exception:
                    pass
            out.append({"id": d.name, "name": name, "tables": ntables, "active": d.name == _active_pid})
    return out


def _init_store():
    """Create the store, migrate any legacy flat layout, open the active project."""
    global _active_pid, _active_name
    STORE.mkdir(exist_ok=True)
    PROJECTS.mkdir(exist_ok=True)
    # migrate the legacy single-workspace layout → projects/default/
    legacy = STORE / "_session.json"
    if legacy.exists() and not (PROJECTS / "default").exists():
        dd = PROJECTS / "default"
        dd.mkdir(parents=True, exist_ok=True)
        legacy.rename(dd / "_session.json")
        for p in STORE.glob("*.parquet"):
            p.rename(dd / p.name)
        try:
            m = json.loads((dd / "_session.json").read_text(encoding="utf-8"))
            m.setdefault("name", "Default")
            (dd / "_session.json").write_text(json.dumps(m), encoding="utf-8")
        except Exception:
            pass
        print("[migrate] moved legacy workspace into projects/default/")
    # choose the active project: last opened, else first on disk, else a fresh default
    pid = None
    af = STORE / "_active.txt"
    if af.exists():
        pid = (af.read_text(encoding="utf-8").strip() or None)
    if not pid or not (PROJECTS / pid).exists():
        dirs = [p.name for p in PROJECTS.iterdir() if p.is_dir()]
        pid = "default" if "default" in dirs else (dirs[0] if dirs else "default")
    (PROJECTS / pid).mkdir(parents=True, exist_ok=True)
    _load_project(pid)
    if not (_proj_dir(pid) / "_session.json").exists():
        if pid == "default":
            _active_name = "Default"
        _persist()  # write an (empty) manifest so the project is listable
    _write_active()


# ──────────────────────────────────────────────────────────────────────────────
# Safe formula evaluator — turns a user expression into a Polars expression.
# Only column refs, literals, arithmetic, comparisons and a whitelist of
# functions are allowed (walked via ast — no Python eval / no attribute access).
# ──────────────────────────────────────────────────────────────────────────────
_FUNCS = {
    "log": lambda a: a.log(), "ln": lambda a: a.log(), "log10": lambda a: a.log10(),
    "log2": lambda a: a.log(2), "log1p": lambda a: (a + 1).log(),
    "sqrt": lambda a: a.sqrt(), "abs": lambda a: a.abs(), "exp": lambda a: a.exp(),
    "floor": lambda a: a.floor(), "ceil": lambda a: a.ceil(), "sign": lambda a: a.sign(),
    "sin": lambda a: a.sin(), "cos": lambda a: a.cos(), "tan": lambda a: a.tan(),
    "round": lambda a, n=0: a.round(int(n)),
    "min": lambda *a: pl.min_horizontal(*a), "max": lambda *a: pl.max_horizontal(*a),
    "pow": lambda a, b: a ** b, "coalesce": lambda *a: pl.coalesce(*a),
    "abs_": lambda a: a.abs(),
}
_BINOPS = {
    ast.Add: lambda a, b: a + b, ast.Sub: lambda a, b: a - b,
    ast.Mult: lambda a, b: a * b, ast.Div: lambda a, b: a / b,
    ast.Mod: lambda a, b: a % b, ast.Pow: lambda a, b: a ** b,
    ast.FloorDiv: lambda a, b: (a / b).floor(),
}
_CMPOPS = {
    ast.Gt: lambda a, b: a > b, ast.GtE: lambda a, b: a >= b,
    ast.Lt: lambda a, b: a < b, ast.LtE: lambda a, b: a <= b,
    ast.Eq: lambda a, b: a == b, ast.NotEq: lambda a, b: a != b,
}


def _as_expr(v):
    return v if isinstance(v, pl.Expr) else pl.lit(v)


def _eval_node(node, cols: set[str]):
    if isinstance(node, ast.Expression):
        return _eval_node(node.body, cols)
    if isinstance(node, ast.BinOp) and type(node.op) in _BINOPS:
        return _BINOPS[type(node.op)](_as_expr(_eval_node(node.left, cols)), _as_expr(_eval_node(node.right, cols)))
    if isinstance(node, ast.UnaryOp):
        v = _eval_node(node.operand, cols)
        if isinstance(node.op, ast.USub):
            return -_as_expr(v)
        if isinstance(node.op, ast.UAdd):
            return _as_expr(v)
        if isinstance(node.op, ast.Not):
            return ~_as_expr(v)
    if isinstance(node, ast.BoolOp):
        parts = [_as_expr(_eval_node(v, cols)) for v in node.values]
        out = parts[0]
        for p in parts[1:]:
            out = (out & p) if isinstance(node.op, ast.And) else (out | p)
        return out
    if isinstance(node, ast.Compare) and len(node.ops) == 1 and type(node.ops[0]) in _CMPOPS:
        return _CMPOPS[type(node.ops[0])](_as_expr(_eval_node(node.left, cols)), _as_expr(_eval_node(node.comparators[0], cols)))
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
        fname = node.func.id
        if fname == "col":  # col('Some Name') → reference a column by literal name
            if len(node.args) == 1 and isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str):
                cn = node.args[0].value
                if cn not in cols:
                    raise ValueError(f"unknown column: {cn}")
                return pl.col(cn)
            raise ValueError("col() takes one quoted column name")
        if fname in _FUNCS:
            args = [_as_expr(_eval_node(a, cols)) for a in node.args]
            return _FUNCS[fname](*args)
        raise ValueError(f"function not allowed: {fname}")
    if isinstance(node, ast.Name):
        if node.id in cols:
            return pl.col(node.id)
        raise ValueError(f"unknown column: {node.id}")
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float, str, bool)):
        return node.value
    raise ValueError("unsupported expression")


def _safe_expr(expr: str, df: pl.DataFrame) -> pl.Expr:
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError as e:
        raise ValueError(f"could not parse formula: {e}")
    return _as_expr(_eval_node(tree, set(df.columns)))


# ──────────────────────────────────────────────────────────────────────────────
# API
# ──────────────────────────────────────────────────────────────────────────────
@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    data = await file.read()
    fname = (file.filename or "table")
    low = fname.lower()
    try:
        if low.endswith((".parquet", ".pq")):
            df = pl.read_parquet(io.BytesIO(data))
        elif low.endswith((".ndjson", ".jsonl")):
            df = pl.read_ndjson(io.BytesIO(data))
        elif low.endswith(".json"):
            df = pl.read_json(io.BytesIO(data))
        else:  # default: CSV / TSV
            sep = "\t" if low.endswith(".tsv") else ","
            df = pl.read_csv(io.BytesIO(data), separator=sep, try_parse_dates=True, infer_schema_length=2000)
    except Exception as e:
        raise HTTPException(400, f"Could not read '{fname}': {e}")
    name = _register(df, fname.rsplit(".", 1)[0], "upload", fname)
    return {"name": name, "rows": df.height, "cols": df.width, "schema": _schema(df)}


@app.get("/api/tables")
def list_tables():
    return [
        {
            "name": n,
            "rows": TABLES[n].height,
            "cols": TABLES[n].width,
            "schema": _schema(TABLES[n]),
            "kind": META.get(n, {}).get("kind", "upload"),
            "detail": META.get(n, {}).get("detail", ""),
            "charts": [{"id": c["id"], "title": c["title"], "cfg": c["cfg"]} for c in CHARTS.get(n, [])],
            "history": len(HISTORY.get(n, [])),
            "undo": (HISTORY.get(n) or [{}])[-1].get("label") if HISTORY.get(n) else None,
            "recipe": len(RECIPE.get(n, [])),
        }
        for n in TABLES
    ]


@app.delete("/api/table/{name}")
def drop_table(name: str):
    TABLES.pop(name, None)
    META.pop(name, None)
    CHARTS.pop(name, None)  # a table's saved charts die with it
    HISTORY.pop(name, None)
    RECIPE.pop(name, None)
    _persist()
    return {"ok": True}


@app.post("/api/table/rename")
async def rename_table(req: Request):
    q = await req.json()
    old, new = q.get("old"), (q.get("new") or "").strip()
    if old not in TABLES:
        raise HTTPException(404, "table not found")
    if not new:
        raise HTTPException(400, "new name is empty")
    if new == old:
        return {"name": old}
    if new in TABLES:
        raise HTTPException(400, f"a table named '{new}' already exists")
    for store in (TABLES, META, CHARTS, HISTORY, RECIPE):  # move every keyed entry
        if old in store:
            store[new] = store.pop(old)
    _persist()
    return {"name": new}


@app.post("/api/charts/save")
async def chart_save(req: Request):
    """Upsert a saved chart config under a table. Omit id to create a new one."""
    q = await req.json()
    name = q.get("table")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    title = (q.get("title") or "Chart").strip() or "Chart"
    cfg = q.get("cfg") or {}
    cid = q.get("id")
    lst = CHARTS.setdefault(name, [])
    existing = next((c for c in lst if c["id"] == cid), None) if cid else None
    if existing:
        existing["title"], existing["cfg"] = title, cfg
    else:
        cid = secrets.token_hex(4)
        lst.append({"id": cid, "title": title, "cfg": cfg})
    _persist()
    return {"id": cid, "title": title}


@app.post("/api/charts/delete")
async def chart_delete(req: Request):
    q = await req.json()
    name, cid = q.get("table"), q.get("id")
    CHARTS[name] = [c for c in CHARTS.get(name, []) if c["id"] != cid]
    _persist()
    return {"ok": True}


@app.post("/api/data")
async def get_data(req: Request):
    q = await req.json()
    name = q.get("table")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    df = _apply_query(TABLES[name], q)
    return _payload(df, q.get("page", 1), int(q.get("page_size", 50)))


@app.post("/api/raw")
async def get_raw(req: Request):
    """Stream the (filtered) rows of a table for the client-side chart engine.

    Returns numeric_cols as integer column indices to match the chart builder's
    expectations. Capped to keep the browser responsive."""
    q = await req.json()
    name = q.get("table")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    df = _apply_query(TABLES[name], q)
    total = df.height                       # true row count BEFORE the responsiveness cap
    cap = int(q.get("limit", 50000))
    if df.height > cap:
        df = df.head(cap)
    cols = df.columns
    numeric_idx = [i for i, c in enumerate(cols) if _is_num(df.schema[c])]
    rows = [[_sanitize(v) for v in r] for r in df.rows()]
    return {"cols": cols, "rows": rows, "numeric_cols": numeric_idx, "total": total}


@app.post("/api/values")
async def col_values(req: Request):
    """Distinct values of one column, for the header value-checklist filter."""
    q = await req.json()
    name, col = q.get("table"), q.get("col")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    df = TABLES[name]
    if col not in df.columns:
        raise HTTPException(400, "column not found")
    is_num = _is_num(df.schema[col])
    s = df.get_column(col)
    try:
        uniq = df.select(pl.col(col).cast(pl.Utf8)).drop_nulls().unique().to_series()
        total = uniq.len()
        vals = sorted(uniq.to_list())
    except Exception:
        vals, total = [], 0
    search = str(q.get("search") or "").strip().lower()
    if search:
        vals = [v for v in vals if search in v.lower()]
    limit = int(q.get("limit", 2000))
    out = {
        "values": vals[:limit],
        "total_unique": int(total),
        "shown": min(len(vals), limit),
        "numeric": is_num,
    }
    if is_num and df.height:
        try:
            out["min"], out["max"] = _sanitize(s.min()), _sanitize(s.max())
        except Exception:
            pass
    return out


@app.post("/api/pivot")
async def pivot(req: Request):
    q = await req.json()
    name = q.get("table")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    df = TABLES[name]
    index = q.get("index") or []
    columns = q.get("columns") or []
    values = q.get("values")
    if isinstance(values, str):
        values = [values]
    values = [v for v in (values or []) if v in df.columns]
    agg = {"avg": "mean", "average": "mean"}.get(q.get("agg", "sum"), q.get("agg", "sum"))
    if not index or not values:
        raise HTTPException(400, "pivot needs at least one Rows field and at least one Values field")
    try:
        if columns:
            try:
                pdf = df.pivot(on=columns, index=index, values=values, aggregate_function=agg)
            except TypeError:  # older Polars used columns=
                pdf = df.pivot(columns=columns, index=index, values=values, aggregate_function=agg)
        else:
            pdf = df.group_by(index, maintain_order=True).agg(
                [_agg_expr(v, agg).alias(f"{agg}_{v}") for v in values]
            )
    except Exception as e:
        raise HTTPException(400, f"Pivot failed: {e}")
    new = _register(pdf, f"{name}__pivot", "pivot", f"{agg}({', '.join(values)}) by {index}×{columns}")
    return {"name": new, **_payload(pdf)}


@app.post("/api/groupby")
async def groupby(req: Request):
    q = await req.json()
    name = q.get("table")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    df = TABLES[name]
    by = q.get("by") or []
    aggs = q.get("aggs") or []
    if not by:
        raise HTTPException(400, "group-by needs at least one key column")
    exprs = []
    seen: set[str] = set()
    for a in aggs:
        col = a.get("col")
        ag = {"avg": "mean"}.get(a.get("agg", "sum"), a.get("agg", "sum"))
        if col in df.columns:
            alias = f"{ag}_{col}"
            base, n = alias, 2
            while alias in seen:  # multiple aggs on the same column → unique names
                alias, n = f"{base}_{n}", n + 1
            seen.add(alias)
            exprs.append(_agg_expr(col, ag).alias(alias))
    if not exprs:
        exprs = [pl.len().alias("count")]
    gdf = df.group_by(by, maintain_order=True).agg(exprs).sort(by)
    detail = f"by {by} · {len(exprs)} metric(s)"
    new = _register(gdf, f"{name}__grouped", "groupby", detail)
    return {"name": new, **_payload(gdf)}


def _do_join(ldf, rdf, left_on, right_on, how, validate, suffix):
    how = "full" if how == "outer" else how
    try:
        if how == "cross":
            return ldf.join(rdf, how="cross", suffix=suffix)
        return ldf.join(rdf, left_on=left_on, right_on=right_on, how=how, validate=validate, suffix=suffix)
    except TypeError:  # very old Polars without validate=
        return ldf.join(rdf, left_on=left_on, right_on=right_on, how=how, suffix=suffix)


def _select_join_cols(jdf, left_df, right_df, left_on, keep_left, keep_right, suffix):
    """Keep only the columns the user asked for. keep_left/keep_right are pre-join
    column names; right columns that clashed with the left are renamed with suffix."""
    if keep_left is None and keep_right is None:
        return jdf
    cols = jdf.columns
    left_set = set(left_df.columns)
    final: list[str] = []

    def _add(c):
        if c in cols and c not in final:
            final.append(c)

    for k in (left_on or []):  # always retain the join keys
        _add(k)
    for c in (keep_left or []):
        _add(c)
    for c in (keep_right or []):
        if c in cols and c not in left_set:
            _add(c)                 # right-only column kept its name
        elif (c + suffix) in cols:
            _add(c + suffix)        # clashed → renamed with suffix
        else:
            _add(c)
    final = [c for c in final if c in cols]
    return jdf.select(final) if final else jdf


@app.post("/api/join")
async def join(req: Request):
    q = await req.json()
    left, right = q.get("left"), q.get("right")
    if left not in TABLES or right not in TABLES:
        raise HTTPException(404, "table not found")
    how = q.get("how", "inner")
    validate = q.get("validate", "m:m")  # 1:1 / 1:m / m:1 / m:m
    left_on = q.get("left_on") or []
    right_on = q.get("right_on") or []
    keep_left = q.get("keep_left")    # None → keep all
    keep_right = q.get("keep_right")  # None → keep all
    suffix = q.get("suffix") or f"_{right}"
    if how != "cross" and (not left_on or not right_on or len(left_on) != len(right_on)):
        raise HTTPException(400, "pick matching key columns on both tables")
    try:
        jdf = _do_join(TABLES[left], TABLES[right], left_on, right_on, how, validate, suffix)
        jdf = _select_join_cols(jdf, TABLES[left], TABLES[right], left_on, keep_left, keep_right, suffix)
    except Exception as e:
        # Polars raises if cardinality validation fails — surface it clearly
        raise HTTPException(400, f"Join failed ({how}, {validate}): {e}")
    detail = (
        f"{left} ⋈ {right} cross"
        if how == "cross"
        else f"{left}.{left_on} {how} {right}.{right_on} [{validate}]"
    )
    new = _register(jdf, f"{left}__{how}__{right}", "join", detail)
    return {"name": new, "detail": detail, **_payload(jdf)}


# ──────────────────────────────────────────────────────────────────────────────
# Cleaning + feature engineering — in-place transforms with undo
# ──────────────────────────────────────────────────────────────────────────────
_DTYPES = {
    "int": pl.Int64, "float": pl.Float64, "str": pl.Utf8, "bool": pl.Boolean,
    "date": pl.Date, "datetime": pl.Datetime,
}


def _do_transform(df: pl.DataFrame, op: str, p: dict) -> tuple[pl.DataFrame, str]:
    """Return (new_df, human label). Raises ValueError on bad params."""
    cols = p.get("cols") or ([p["col"]] if p.get("col") else [])

    # ── cleaning ──
    if op == "fillna":
        method, val = p.get("method", "constant"), p.get("value")
        exprs = []
        for c in cols:
            e = pl.col(c)
            if method == "mean":
                e = e.fill_null(pl.col(c).mean())
            elif method == "median":
                e = e.fill_null(pl.col(c).median())
            elif method == "mode":
                e = e.fill_null(pl.col(c).mode().first())
            elif method == "ffill":
                e = e.forward_fill()
            elif method == "bfill":
                e = e.backward_fill()
            else:  # constant
                cast = float(val) if _is_num(df.schema[c]) and val not in (None, "") else val
                e = e.fill_null(cast)
            exprs.append(e.alias(c))
        return df.with_columns(exprs), f"fill nulls ({method}) in {cols}"

    if op == "dropna":
        if p.get("axis") == "cols":  # drop columns that are entirely null
            keep = [c for c in df.columns if df[c].null_count() < df.height]
            return df.select(keep), "drop all-null columns"
        sub = cols or None
        return df.drop_nulls(subset=sub), f"drop rows with nulls{(' in '+str(cols)) if cols else ''}"

    if op == "dropdup":
        sub = cols or None
        return df.unique(subset=sub, keep="first", maintain_order=True), f"drop duplicate rows{(' on '+str(cols)) if cols else ''}"

    if op == "cast":
        c, dt, fmt = p["col"], p.get("dtype", "str"), p.get("fmt")
        if dt in ("date", "datetime"):
            tgt = pl.Date if dt == "date" else pl.Datetime
            s = pl.col(c).cast(pl.Utf8)
            e = s.str.strptime(tgt, format=fmt, strict=False) if fmt else (s.str.to_date(strict=False) if dt == "date" else s.str.to_datetime(strict=False))
        else:
            e = pl.col(c).cast(_DTYPES[dt], strict=False)
        return df.with_columns(e.alias(c)), f"cast {c} → {dt}"

    if op == "strclean":
        sub = cols or [c for c in df.columns if df.schema[c] == pl.Utf8]
        actions = p.get("ops") or ["trim"]
        exprs = []
        for c in sub:
            e = pl.col(c).cast(pl.Utf8)
            for a in actions:
                if a == "trim":
                    e = e.str.strip_chars()
                elif a == "lower":
                    e = e.str.to_lowercase()
                elif a == "upper":
                    e = e.str.to_uppercase()
                elif a == "title":
                    e = e.str.to_titlecase()
                elif a == "collapse":
                    e = e.str.replace_all(r"\s+", " ")
            exprs.append(e.alias(c))
        return df.with_columns(exprs), f"clean text {actions} in {sub}"

    if op == "replace":
        c, find, repl = p["col"], p.get("find", ""), p.get("repl", "")
        if p.get("regex"):
            e = pl.col(c).cast(pl.Utf8).str.replace_all(find, repl)
        else:
            e = pl.col(c).cast(pl.Utf8).str.replace_all(find, repl, literal=True)
        return df.with_columns(e.alias(c)), f"replace '{find}'→'{repl}' in {c}"

    if op == "outlier":
        c, method = p["col"], p.get("method", "iqr")
        s = df[c]
        if method == "zscore":
            k = float(p.get("k", 3))
            mu, sd = s.mean(), s.std()
            lo, hi = mu - k * sd, mu + k * sd
        elif method == "manual":
            lo, hi = p.get("lo"), p.get("hi")
            lo = float(lo) if lo not in (None, "") else s.min()
            hi = float(hi) if hi not in (None, "") else s.max()
        else:  # iqr
            k = float(p.get("k", 1.5))
            q1, q3 = s.quantile(0.25), s.quantile(0.75)
            iqr = (q3 - q1) if q1 is not None and q3 is not None else 0
            lo, hi = q1 - k * iqr, q3 + k * iqr
        if p.get("drop"):
            return df.filter((pl.col(c) >= lo) & (pl.col(c) <= hi)), f"drop outliers in {c} ({method})"
        return df.with_columns(pl.col(c).clip(lo, hi).alias(c)), f"clip outliers in {c} ({method})"

    if op == "rename":
        return df.rename({p["col"]: p["to"]}), f"rename {p['col']} → {p['to']}"

    if op == "dropcols":
        return df.drop([c for c in cols if c in df.columns]), f"drop columns {cols}"

    if op == "reorder":
        order = [c for c in (p.get("order") or []) if c in df.columns]
        order += [c for c in df.columns if c not in order]
        return df.select(order), "reorder columns"

    # ── feature engineering ──
    if op == "compute":
        name = (p.get("name") or "feature").strip() or "feature"
        expr = _safe_expr(p.get("expr", ""), df)
        return df.with_columns(expr.alias(name)), f"compute {name} = {p.get('expr','')}"

    if op == "bin":
        c, method, newname = p["col"], p.get("method", "quantile"), (p.get("newname") or (p["col"] + "_bin"))
        nbins = int(p.get("bins", 4))
        labels = p.get("labels") or None
        s = df[c]
        if method == "custom":
            edges = [float(x) for x in (p.get("edges") or [])]
            e = pl.col(c).cut(edges, labels=labels)
        elif method == "width":
            lo, hi = s.min(), s.max()
            step = (hi - lo) / nbins if hi != lo else 1
            breaks = [lo + step * i for i in range(1, nbins)]
            e = pl.col(c).cut(breaks, labels=labels)
        else:  # quantile
            qs = [i / nbins for i in range(1, nbins)]
            e = pl.col(c).qcut(qs, labels=labels, allow_duplicates=True)
        return df.with_columns(e.cast(pl.Utf8).alias(newname)), f"bin {c} → {newname} ({method},{nbins})"

    if op == "encode":
        c, method = p["col"], p.get("method", "onehot")
        if method == "label":
            e = pl.col(c).cast(pl.Utf8).cast(pl.Categorical).to_physical().alias(p.get("newname") or (c + "_code"))
            return df.with_columns(e), f"label-encode {c}"
        try:
            out = df.to_dummies(columns=[c], drop_first=bool(p.get("drop_first")))
        except TypeError:
            out = df.to_dummies(columns=[c])
        return out, f"one-hot encode {c}"

    if op == "dateparts":
        c, parts = p["col"], p.get("parts") or ["year", "month"]
        s = pl.col(c)
        m = {
            "year": s.dt.year(), "month": s.dt.month(), "day": s.dt.day(),
            "weekday": s.dt.weekday(), "quarter": s.dt.quarter(),
            "week": s.dt.week(), "hour": s.dt.hour(), "ordinal_day": s.dt.ordinal_day(),
        }
        exprs = [m[k].alias(f"{c}_{k}") for k in parts if k in m]
        return df.with_columns(exprs), f"date parts {parts} from {c}"

    if op == "window":
        func, c = p.get("func", "cumsum"), p["col"]
        by = p.get("by") or []
        order = p.get("order")
        newname = p.get("newname") or f"{c}_{func}"
        w = int(p.get("window", 3))
        n = int(p.get("n", 1))
        d = df.sort(order) if order and order in df.columns else df
        e = pl.col(c)
        if func == "rank":
            e = e.rank(method="ordinal")
        elif func == "cumsum":
            e = e.cum_sum()
        elif func == "rolling_mean":
            e = e.rolling_mean(window_size=w)
        elif func == "lag":
            e = e.shift(n)
        elif func == "lead":
            e = e.shift(-n)
        elif func == "pct_change":
            e = e.pct_change()
        if by:
            e = e.over(by)
        return d.with_columns(e.alias(newname)), f"window {func} on {c}"

    if op == "scale":
        method = p.get("method", "zscore")
        exprs = []
        for c in cols:
            if method == "minmax":
                lo, hi = pl.col(c).min(), pl.col(c).max()
                e = (pl.col(c) - lo) / (hi - lo)
            else:  # zscore
                e = (pl.col(c) - pl.col(c).mean()) / pl.col(c).std()
            nm = c if p.get("inplace") else f"{c}_scaled"
            exprs.append(e.alias(nm))
        return df.with_columns(exprs), f"scale ({method}) {cols}"

    if op == "split":
        c = p["col"]
        if c not in df.columns:
            raise ValueError(f"no such column: {c}")
        sep = p.get("sep", "")
        if sep == "":
            raise ValueError("a delimiter / character is required")
        n = max(2, int(p.get("parts", 2)))
        # target names: use supplied names, padded with col_1, col_2, … and trimmed to n
        names = [str(x).strip() for x in (p.get("into") or []) if str(x).strip()]
        names = (names + [f"{c}_{i + 1}" for i in range(len(names), n)])[:n]
        src = pl.col(c).cast(pl.Utf8)
        if p.get("remainder", True):
            # last column keeps everything after the (n-1)th delimiter  ("a-b-c" → "a","b-c")
            st = src.str.splitn(sep, n)
            exprs = [st.struct.field(f"field_{i}").alias(names[i]) for i in range(n)]
        else:
            # split on every delimiter, keep only the first n parts (rest discarded)
            lst = src.str.split(sep)
            exprs = [lst.list.get(i, null_on_oob=True).alias(names[i]) for i in range(n)]
        out = df.with_columns(exprs)
        if p.get("drop"):
            out = out.drop(c)
        return out, f"split {c} by '{sep}' → {names}"

    if op == "explode":
        c = p["col"]
        if c not in df.columns:
            raise ValueError(f"no such column: {c}")
        sep = p.get("sep", "")
        is_list = isinstance(df.schema[c], pl.List)
        out = df
        if sep:
            e = pl.col(c).cast(pl.Utf8).str.split(sep)
            if p.get("strip", True):
                e = e.list.eval(pl.element().str.strip_chars())
            out = out.with_columns(e.alias(c))
        elif not is_list:
            raise ValueError("a delimiter / character is required to split before exploding")
        out = out.explode(c)  # one part per row, other columns duplicated
        if p.get("dropna", True):
            out = out.filter(pl.col(c).is_not_null() & (pl.col(c).cast(pl.Utf8) != ""))
        return out, f"explode {c}" + (f" by '{sep}'" if sep else "")

    raise ValueError(f"unknown op: {op}")


@app.post("/api/transform")
async def transform(req: Request):
    q = await req.json()
    name, op = q.get("table"), q.get("op")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    params = q.get("params") or {}
    try:
        new_df, label = _do_transform(TABLES[name], op, params)
    except Exception as e:
        raise HTTPException(400, f"{op} failed: {e}")
    _commit(name, new_df, label)
    RECIPE.setdefault(name, []).append({"op": op, "params": params, "label": label})  # replayable
    _persist()
    return {"label": label, "history": len(HISTORY.get(name, [])), **_payload(new_df, 1, int(q.get("page_size", 50)))}


@app.post("/api/undo")
async def undo(req: Request):
    q = await req.json()
    name = q.get("table")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    hist = HISTORY.get(name) or []
    if not hist:
        raise HTTPException(400, "nothing to undo")
    prev = hist.pop()
    TABLES[name] = prev["df"]
    if RECIPE.get(name):
        RECIPE[name].pop()  # keep the recipe in lock-step with the undo stack
    _persist()
    return {"undone": prev["label"], "history": len(hist), **_payload(TABLES[name], 1, int(q.get("page_size", 50)))}


@app.get("/api/recipe/{name}")
def get_recipe(name: str):
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    return {"steps": [{"op": s["op"], "label": s["label"]} for s in RECIPE.get(name, [])]}


@app.post("/api/recipe/apply")
async def apply_recipe(req: Request):
    """Replay one table's recorded clean/feature steps onto another table."""
    q = await req.json()
    source, target = q.get("source"), q.get("target")
    if source not in TABLES or target not in TABLES:
        raise HTTPException(404, "table not found")
    steps = RECIPE.get(source, [])
    if not steps:
        raise HTTPException(400, f"'{source}' has no recorded steps to apply")
    applied = 0
    for i, s in enumerate(steps):
        try:
            new_df, label = _do_transform(TABLES[target], s["op"], s.get("params") or {})
        except Exception as e:
            raise HTTPException(400, f"step {i + 1} ({s['op']}) failed on '{target}': {e}")
        _commit(target, new_df, label)
        RECIPE.setdefault(target, []).append({"op": s["op"], "params": s.get("params") or {}, "label": label})
        applied += 1
    _persist()
    return {"applied": applied, **_payload(TABLES[target], 1, int(q.get("page_size", 50)))}


@app.post("/api/union")
async def union(req: Request):
    """Stack (concat) two or more tables. 'diagonal' tolerates differing columns."""
    q = await req.json()
    names = [n for n in (q.get("tables") or []) if n in TABLES]
    if len(names) < 2:
        raise HTTPException(400, "pick at least two tables to stack")
    how = q.get("how", "diagonal")
    try:
        if how == "vertical":
            udf = pl.concat([TABLES[n] for n in names], how="vertical_relaxed")
        else:
            udf = pl.concat([TABLES[n] for n in names], how="diagonal_relaxed")
    except Exception as e:
        raise HTTPException(400, f"Union failed: {e}")
    new = _register(udf, "__".join(names[:2]) + "__union", "union", f"{how} concat of {names}")
    return {"name": new, **_payload(udf)}


@app.post("/api/profile")
async def profile(req: Request):
    q = await req.json()
    name = q.get("table")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    df = TABLES[name]
    n = df.height
    out = []
    for c in df.columns:
        s = df[c]
        nulls = s.null_count()
        uniq = s.n_unique()
        col = {
            "name": c, "dtype": str(df.schema[c]), "numeric": _is_num(df.schema[c]),
            "nulls": nulls, "null_pct": round(100 * nulls / n, 1) if n else 0,
            "unique": uniq,
            "constant": uniq <= 1,
            "high_card": (not _is_num(df.schema[c])) and n > 0 and uniq > 0.5 * n and uniq > 50,
        }
        if _is_num(df.schema[c]) and n:
            try:
                col.update({
                    "min": _sanitize(s.min()), "max": _sanitize(s.max()),
                    "mean": _sanitize(s.mean()), "median": _sanitize(s.median()),
                    "std": _sanitize(s.std()),
                })
            except Exception:
                pass
        try:
            vc = s.drop_nulls().value_counts(sort=True).head(5)
            col["top"] = [[_sanitize(r[0]), int(r[1])] for r in vc.rows()]
        except Exception:
            col["top"] = []
        out.append(col)
    dup = n - df.unique().height
    return {"rows": n, "cols": df.width, "duplicates": dup, "columns": out}


@app.post("/api/corr")
async def corr(req: Request):
    q = await req.json()
    name = q.get("table")
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    df = TABLES[name]
    nums = _numeric_cols(df)
    if len(nums) < 2:
        raise HTTPException(400, "need at least 2 numeric columns for correlation")
    try:
        cm = df.select(nums).corr()
        matrix = [[_sanitize(v) for v in row] for row in cm.rows()]
    except Exception as e:
        raise HTTPException(400, f"correlation failed: {e}")
    return {"cols": nums, "matrix": matrix}


@app.post("/api/session/clear")
def session_clear():
    """Empty the CURRENTLY-OPEN project (keep the project itself)."""
    TABLES.clear(); META.clear(); CHARTS.clear(); HISTORY.clear(); RECIPE.clear()
    _persist()  # rewrites an empty manifest and prunes the project's parquet files
    return {"ok": True}


# ──────────────────────────────────────────────────────────────────────────────
# Projects — each is an isolated workspace (its own tables/charts/recipes on disk)
# ──────────────────────────────────────────────────────────────────────────────
@app.get("/api/projects")
def projects_list():
    return {"projects": _list_projects(), "active": _active_pid, "active_name": _active_name}


@app.post("/api/projects/create")
async def projects_create(req: Request):
    global _active_pid, _active_name
    q = await req.json()
    name = (q.get("name") or "").strip() or "Untitled"
    pid = _unique_pid(name)
    (PROJECTS / pid).mkdir(parents=True, exist_ok=True)
    TABLES.clear(); META.clear(); CHARTS.clear(); HISTORY.clear(); RECIPE.clear()
    _active_pid, _active_name = pid, name
    _persist()
    return {"ok": True, "id": pid, "name": name, "active": _active_pid,
            "active_name": _active_name, "projects": _list_projects()}


@app.post("/api/projects/open")
async def projects_open(req: Request):
    q = await req.json()
    pid = q.get("id")
    if not pid or not (PROJECTS / pid).exists():
        raise HTTPException(404, "project not found")
    _load_project(pid)
    _write_active()
    return {"ok": True, "id": pid, "active": _active_pid,
            "active_name": _active_name, "projects": _list_projects()}


@app.post("/api/projects/rename")
async def projects_rename(req: Request):
    global _active_name
    q = await req.json()
    pid = q.get("id")
    name = (q.get("name") or "").strip()
    if not pid or not (PROJECTS / pid).exists():
        raise HTTPException(404, "project not found")
    if not name:
        raise HTTPException(400, "name required")
    if pid == _active_pid:
        _active_name = name
        _persist()
    else:
        sess = PROJECTS / pid / "_session.json"
        m = {}
        if sess.exists():
            try:
                m = json.loads(sess.read_text(encoding="utf-8"))
            except Exception:
                m = {}
        m["name"] = name
        m.setdefault("tables", []); m.setdefault("charts", {}); m.setdefault("recipe", {})
        sess.write_text(json.dumps(m), encoding="utf-8")
    return {"ok": True, "active": _active_pid, "active_name": _active_name, "projects": _list_projects()}


@app.post("/api/projects/delete")
async def projects_delete(req: Request):
    global _active_pid, _active_name
    q = await req.json()
    pid = q.get("id")
    if not pid or not (PROJECTS / pid).exists():
        raise HTTPException(404, "project not found")
    shutil.rmtree(PROJECTS / pid, ignore_errors=True)
    if pid == _active_pid:
        dirs = [p.name for p in PROJECTS.iterdir() if p.is_dir()]
        if dirs:
            _load_project(dirs[0])
        else:
            (PROJECTS / "default").mkdir(parents=True, exist_ok=True)
            _load_project("default")
            _active_name = "Default"
            _persist()
        _write_active()
    return {"ok": True, "active": _active_pid, "active_name": _active_name, "projects": _list_projects()}


@app.get("/api/export/{name}")
def export_table(name: str, fmt: str = "csv"):
    if name not in TABLES:
        raise HTTPException(404, "table not found")
    df = TABLES[name]
    fmt = (fmt or "csv").lower()
    try:
        if fmt == "parquet":
            buf = io.BytesIO(); df.write_parquet(buf)
            data, mime, ext = buf.getvalue(), "application/octet-stream", "parquet"
        elif fmt == "json":
            data, mime, ext = df.write_json().encode(), "application/json", "json"
        elif fmt in ("xlsx", "excel"):
            buf = io.BytesIO()
            try:
                df.write_excel(buf)  # needs xlsxwriter
            except Exception as e:
                raise HTTPException(400, f"Excel export needs the 'xlsxwriter' package: {e}")
            data, mime, ext = buf.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"
        else:
            data, mime, ext = df.write_csv().encode(), "text/csv", "csv"
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, f"export failed: {e}")
    return Response(
        data, media_type=mime,
        headers={"Content-Disposition": f'attachment; filename="{name}.{ext}"'},
    )


@app.get("/", response_class=HTMLResponse)
def index():
    return FRONTEND_HTML


# ──────────────────────────────────────────────────────────────────────────────
# Frontend (single page; talks to the API above). Plain string — not an f-string,
# so JS braces/${} need no escaping.
# ──────────────────────────────────────────────────────────────────────────────
FRONTEND_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>pudbo · Polars Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{--bg:#0d1117;--surface:#161b22;--surface2:#1c2230;--border:#2b3340;--text1:#e6edf3;--text2:#9aa7b4;--text3:#6b7681;--accent:#4f6ef7;--good:#22d3a5;--bad:#f75a7a;
/* aliases so the ported db_fw12 chart CSS resolves against this theme */
--nav-border:var(--border);--toggle-bg:var(--surface2);--input-bg:var(--surface2);--tab-bg:rgba(79,110,247,.16);--accent-text:var(--accent);--tbl-border:var(--border);}
/* ── THEMES (swap the core vars; aliases follow automatically) ── */
body.theme-carbon{--bg:#0b0c0e;--surface:#15171a;--surface2:#1d2024;--border:#2a2e34;--text1:#e8eaed;--text2:#9aa0a6;--text3:#6b7178;--accent:#7c93ff;--good:#34d399;--bad:#f87171;}
body.theme-midnight{--bg:#0a1413;--surface:#10201d;--surface2:#16302b;--border:#214039;--text1:#e3f1ec;--text2:#93b3a9;--text3:#5f7d74;--accent:#2dd4bf;--good:#34d399;--bad:#fb7185;}
body.theme-crimson{--bg:#120b0e;--surface:#1c1014;--surface2:#26161c;--border:#3a222a;--text1:#f4e8ec;--text2:#c096a3;--text3:#8a626d;--accent:#f43f5e;--good:#34d399;--bad:#fb7185;}
body.theme-grape{--bg:#0f0b16;--surface:#191222;--surface2:#231a30;--border:#352847;--text1:#ece6f6;--text2:#a99ac0;--text3:#766a8c;--accent:#a78bfa;--good:#34d399;--bad:#fb7185;}
body.theme-slate{--bg:#f1f5f9;--surface:#ffffff;--surface2:#eef2f7;--border:#d4dde8;--text1:#1e293b;--text2:#52617a;--text3:#8493a8;--accent:#3b6ef5;--good:#0d9b73;--bad:#dc2647;--tab-bg:rgba(59,110,245,.12);}
*{box-sizing:border-box;}
body{margin:0;font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text1);font-size:13px;}
.app{display:flex;height:100vh;}
.sidebar{width:270px;min-width:270px;background:var(--surface);border-right:1px solid var(--border);display:flex;flex-direction:column;overflow:hidden;transition:width .15s,min-width .15s;}
.app.sidebar-collapsed .sidebar{width:0;min-width:0;border-right:none;}
.sb-toggle{font-size:15px;line-height:1;padding:5px 10px;}
.sidebar h2{font-size:13px;margin:0;padding:14px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px;}
.up{padding:12px 16px;border-bottom:1px solid var(--border);}
.up label{display:block;border:1.5px dashed var(--border);border-radius:8px;padding:14px;text-align:center;color:var(--text2);cursor:pointer;font-size:12px;}
.up label:hover{border-color:var(--accent);color:var(--text1);}
.up input{display:none;}
.tlist{flex:1;overflow-y:auto;padding:8px;}
.titem{padding:8px 10px;border-radius:7px;cursor:pointer;margin-bottom:4px;border:1px solid transparent;}
.titem:hover{background:var(--surface2);}
.titem.active{background:var(--surface2);border-color:var(--accent);}
.titem .nm{font-weight:600;font-size:12.5px;word-break:break-all;}
.titem .mt{font-size:10.5px;color:var(--text3);margin-top:2px;}
.titem .kind{font-size:9px;text-transform:uppercase;letter-spacing:.5px;color:var(--accent);}
.charts-nav{margin:-2px 0 6px 14px;border-left:1px solid var(--border);padding-left:6px;}
.chart-nav-link{display:flex;align-items:center;gap:6px;padding:3px 8px;border-radius:6px;cursor:pointer;font-size:11px;color:var(--text2);}
.chart-nav-link:hover{background:var(--surface2);color:var(--text1);}
.chart-nav-link::before{content:'📈';font-size:10px;flex-shrink:0;}
.chart-nav-link .cnl-t{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.chart-nav-link .cnl-x{color:var(--text3);opacity:0;}
.chart-nav-link:hover .cnl-x{opacity:1;}
.chart-nav-link .cnl-x:hover{color:var(--bad);}
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;}
.scroll{flex:1;overflow:auto;}
.toolbar{display:flex;gap:8px;align-items:center;padding:10px 16px;border-bottom:1px solid var(--border);background:var(--surface);flex-wrap:wrap;}
#curName{background:transparent;border:1px solid transparent;border-radius:7px;padding:5px 8px;font-size:14px;font-weight:700;color:var(--text1);min-width:120px;max-width:280px;}
#curName:hover{border-color:var(--border);}
#curName:focus{border-color:var(--accent);background:var(--surface2);outline:none;}
.spinner{display:none;width:15px;height:15px;border:2px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite;flex-shrink:0;}
@keyframes spin{to{transform:rotate(360deg);}}
.chartpanel{padding:0 16px 24px;}
.chartpanel-hdr{font-size:13px;font-weight:700;color:var(--text1);margin:6px 0 10px;display:flex;align-items:center;gap:8px;}
.btn{background:var(--surface2);color:var(--text1);border:1px solid var(--border);border-radius:7px;padding:6px 12px;cursor:pointer;font-size:12px;}
.btn:hover{border-color:var(--accent);}
.btn.accent{background:var(--accent);border-color:var(--accent);color:#fff;}
.btn.good{color:var(--good);border-color:var(--good);}
.seg{display:inline-flex;border:1px solid var(--border);border-radius:7px;overflow:hidden;}
.seg button{background:var(--surface2);border:none;color:var(--text2);padding:6px 10px;cursor:pointer;font-size:12px;}
.seg button.on{background:var(--accent);color:#fff;}
.content{padding:16px;}
.muted{color:var(--text3);}
input,select{background:var(--surface2);color:var(--text1);border:1px solid var(--border);border-radius:6px;padding:6px 8px;font-size:12px;}
table{border-collapse:collapse;width:100%;font-size:12px;}
th,td{border-bottom:1px solid var(--border);padding:6px 10px;text-align:left;white-space:nowrap;}
th{position:sticky;top:0;background:var(--surface);cursor:pointer;user-select:none;}
th .ar{color:var(--accent);font-size:10px;}
th{white-space:nowrap;}
.th-lbl{cursor:pointer;}
.th-menu{cursor:pointer;color:var(--text3);margin-left:6px;font-size:10px;opacity:0;}
th:hover .th-menu{opacity:.7;}
.th-menu:hover{color:var(--accent);opacity:1;}
.th-menu.on{opacity:1;color:var(--accent);}
.th-menu.on::after{content:'';display:inline-block;width:5px;height:5px;border-radius:50%;background:var(--accent);margin-left:3px;vertical-align:middle;}
/* ── column filter dropdown ── */
.colfilter{position:fixed;z-index:9999;background:var(--surface);border:1px solid var(--border);border-radius:10px;box-shadow:0 10px 30px rgba(0,0,0,.45);width:280px;overflow:hidden;display:flex;flex-direction:column;}
.cf-head{display:flex;align-items:baseline;justify-content:space-between;gap:8px;padding:10px 12px 6px;}
.cf-title{font-weight:700;font-size:13px;color:var(--text1);word-break:break-all;}
.cf-sub{font-size:10px;color:var(--text3);white-space:nowrap;}
.cf-actions{display:flex;gap:4px;padding:0 10px 8px;border-bottom:1px solid var(--border);}
.cf-act{background:var(--surface2);border:1px solid var(--border);border-radius:5px;color:var(--text2);font-size:11px;padding:2px 7px;cursor:pointer;}
.cf-act:hover{border-color:var(--accent);color:var(--text1);}
.cf-search{margin:8px 10px 6px;padding:6px 9px;border-radius:6px;border:1px solid var(--border);background:var(--surface2);color:var(--text1);font-size:12px;outline:none;}
.cf-search:focus{border-color:var(--accent);}
.cf-body{max-height:230px;overflow-y:auto;padding:2px 4px;}
.cf-item{display:flex;align-items:center;gap:8px;padding:4px 8px;font-size:12px;color:var(--text2);cursor:pointer;border-radius:5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.cf-item:hover{background:var(--surface2);}
.cf-item input{accent-color:var(--accent);cursor:pointer;flex-shrink:0;}
.cf-item.cf-all{color:var(--text1);border-bottom:1px solid var(--border);border-radius:0;}
.cf-range{display:flex;align-items:center;gap:6px;padding:10px 12px;flex-wrap:wrap;}
.cf-range label{font-size:11px;color:var(--text2);}
.cf-range input{width:80px;background:var(--surface2);border:1px solid var(--border);border-radius:5px;padding:5px 8px;color:var(--text1);font-size:12px;outline:none;}
.cf-range input:focus{border-color:var(--accent);}
.cf-note{font-size:10px;color:var(--text3);padding:2px 12px 4px;}
.cf-foot{display:flex;gap:8px;justify-content:flex-end;padding:8px 12px;border-top:1px solid var(--border);}
.cf-clear{background:var(--surface2);border:1px solid var(--border);color:var(--text2);border-radius:6px;padding:5px 14px;font-size:12px;cursor:pointer;}
.cf-apply{background:var(--accent);border:1px solid var(--accent);color:#fff;border-radius:6px;padding:5px 16px;font-size:12px;cursor:pointer;font-weight:600;}
.cf-clear:hover{color:var(--text1);} .cf-apply:hover{opacity:.9;}
td.num,th.num{text-align:right;font-variant-numeric:tabular-nums;}
tr.sum td{background:var(--surface2);font-weight:700;border-top:2px solid var(--border);position:sticky;bottom:0;}
.pager{display:flex;gap:8px;align-items:center;padding:10px 0;}
.empty{display:flex;min-height:55vh;align-items:center;justify-content:center;color:var(--text3);flex-direction:column;gap:8px;}
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.6);display:none;align-items:flex-start;justify-content:center;padding:40px 16px;overflow:auto;z-index:100;}
.overlay.open{display:flex;}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:20px;width:min(760px,100%);}
.modal h3{margin:0 0 14px;}
.field{display:flex;flex-direction:column;gap:4px;margin-bottom:10px;}
.field label{font-size:11px;color:var(--text2);text-transform:uppercase;letter-spacing:.4px;}
.row{display:flex;gap:12px;flex-wrap:wrap;}
.row .field{flex:1;min-width:150px;}
select[multiple]{min-height:110px;}
.chip{font-size:10px;background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:1px 7px;color:var(--text2);}
.toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:var(--surface2);border:1px solid var(--accent);padding:10px 16px;border-radius:8px;display:none;z-index:200;}
.toast.bad{border-color:var(--bad);color:var(--bad);}
.subhdr{font-size:11px;color:var(--text2);text-transform:uppercase;letter-spacing:.4px;margin:6px 0 4px;}
.minibtn{background:var(--surface2);color:var(--text2);border:1px solid var(--border);border-radius:6px;padding:4px 9px;cursor:pointer;font-size:12px;}
.minibtn:hover{border-color:var(--accent);color:var(--text1);}

/* ===== ported db_fw12 chart-builder CSS ===== */
.chart-builder{background:var(--surface);border:0.5px solid var(--tbl-border);border-radius:10px;font-family:'DM Sans',sans-serif;}
.chart-builder-header{padding:10px 16px;border-bottom:0.5px solid var(--tbl-border);display:flex;align-items:center;justify-content:space-between;}
.chart-builder-title{font-size:12px;font-weight:600;color:var(--text1);display:flex;align-items:center;gap:6px;}
.chart-builder-body{padding:16px;}
.chart-builder-body.collapsed{display:none;}
.chart-edit-btn{padding:3px 10px;border-radius:5px;font-size:11px;cursor:pointer;border:0.5px solid var(--accent);background:var(--tab-bg);color:var(--accent);font-family:'DM Sans',sans-serif;transition:all 0.15s;}
.chart-edit-btn:hover{background:var(--accent);color:#fff;}
.chart-type-row{display:flex;align-items:center;gap:6px;margin-bottom:14px;flex-wrap:wrap;}
.chart-type-label{font-size:11px;color:var(--text3);font-weight:500;text-transform:uppercase;letter-spacing:0.5px;margin-right:4px;}
.chart-type-btn{padding:5px 12px;border-radius:6px;font-size:11px;font-weight:500;cursor:pointer;border:0.5px solid var(--nav-border);background:var(--toggle-bg);color:var(--text2);font-family:'DM Sans',sans-serif;transition:all 0.15s;}
.chart-type-btn:hover{color:var(--text1);border-color:var(--accent);}
.chart-type-btn.active{background:var(--accent);color:#fff;border-color:var(--accent);}
.chart-axes-row{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:14px;}
.chart-axis-group{display:flex;flex-direction:column;gap:6px;}
.chart-axis-label{font-size:11px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;}
.chart-axis-select{background:var(--input-bg);border:0.5px solid var(--nav-border);border-radius:6px;padding:6px 10px;font-size:12px;color:var(--text1);font-family:'DM Sans',sans-serif;outline:none;width:100%;}
.chart-axis-select:focus{border-color:var(--accent);}
.chart-y-checks{display:flex;flex-direction:column;gap:4px;max-height:120px;overflow-y:auto;background:var(--input-bg);border:0.5px solid var(--nav-border);border-radius:6px;padding:6px 10px;}
.chart-y-check{display:flex;align-items:center;gap:7px;font-size:12px;color:var(--text2);cursor:pointer;padding:2px 0;}
.chart-y-check:hover{color:var(--text1);}
.chart-y-check input{accent-color:var(--accent);cursor:pointer;}
.chart-generate-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}
.chart-generate-btn{padding:7px 20px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:none;background:var(--accent);color:#fff;font-family:'DM Sans',sans-serif;transition:all 0.15s;}
.chart-generate-btn:hover{opacity:0.85;}
.chart-output{margin-top:12px;border-radius:8px;overflow:hidden;border:0.5px solid var(--tbl-border);background:var(--bg);min-height:400px;}
.chart-placeholder{display:flex;align-items:center;justify-content:center;height:400px;color:var(--text3);font-size:13px;}
.chart-error{color:#f75a7a;font-size:12px;padding:8px 12px;background:rgba(247,90,122,0.1);border-radius:6px;}
.chart-label-toggle{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--text2);cursor:pointer;padding:5px 10px;border-radius:6px;border:0.5px solid var(--nav-border);background:var(--toggle-bg);font-family:'DM Sans',sans-serif;user-select:none;}
.chart-label-toggle.active{background:var(--tab-bg);color:var(--accent-text);border-color:var(--accent);}
.chart-title-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;}
.chart-title-input{background:var(--input-bg);border:0.5px solid var(--nav-border);border-radius:6px;padding:7px 10px;font-size:12px;color:var(--text1);font-family:'DM Sans',sans-serif;outline:none;width:100%;transition:border-color 0.15s;}
.chart-title-input:focus{border-color:var(--accent);}
.chart-title-input::placeholder{color:var(--text3);}
.chart-title-label{font-size:11px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;}
.chart-axis-label-row{display:flex;gap:10px;margin-bottom:12px;flex-wrap:wrap;}
.chart-axis-label-group{display:flex;flex-direction:column;flex:1;min-width:120px;}
.chart-axis-label-tag{font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:3px;}
.chart-axis-label-input{background:var(--input-bg);border:0.5px solid var(--nav-border);border-radius:5px;padding:5px 9px;font-size:11px;color:var(--text1);font-family:'DM Sans',sans-serif;outline:none;width:100%;box-sizing:border-box;transition:border-color 0.15s;}
.chart-axis-label-input:focus{border-color:var(--accent);}
.chart-axis-label-input::placeholder{color:var(--text3);font-style:italic;}
.chart-rendered-title{font-size:14px;font-weight:600;color:var(--text1);margin:12px 0 4px;letter-spacing:-0.2px;}
.chart-rendered-desc{font-size:12px;color:var(--text2);margin-bottom:8px;line-height:1.5;}
.add-chart-btn{margin-top:10px;width:100%;padding:8px;border-radius:8px;border:1px dashed var(--tbl-border);background:transparent;color:var(--text3);font-size:12px;font-weight:500;cursor:pointer;font-family:'DM Sans',sans-serif;transition:all 0.15s;}
.add-chart-btn:hover{color:var(--accent);border-color:var(--accent);background:var(--tab-bg);}
.chart-instance{position:relative;}
.chart-instance-remove{position:absolute;top:10px;right:10px;padding:3px 10px;border-radius:5px;font-size:11px;cursor:pointer;border:0.5px solid var(--nav-border);background:var(--surface);color:var(--text3);font-family:'DM Sans',sans-serif;z-index:10;}
.chart-instance-remove:hover{color:#f75a7a;border-color:#f75a7a;}
.charts-container{margin-top:8px;}

/* ── CHART BUILDER v2 CONTROLS ── */
.chart-options-row{display:flex;align-items:center;gap:6px;margin-bottom:12px;flex-wrap:wrap;}
.chart-options-row .chart-type-label{margin-right:2px;}
.chart-opt-group{display:flex;align-items:center;gap:4px;background:var(--toggle-bg);border:0.5px solid var(--nav-border);border-radius:6px;overflow:hidden;}
.chart-opt-btn{padding:4px 10px;font-size:11px;font-weight:500;cursor:pointer;color:var(--text3);background:transparent;border:none;font-family:'DM Sans',sans-serif;transition:all 0.15s;white-space:nowrap;}
.chart-opt-btn:hover{color:var(--text2);}
.chart-opt-btn.active{background:var(--tab-bg);color:var(--accent-text);}
.chart-opt-select{background:var(--input-bg);border:0.5px solid var(--nav-border);border-radius:5px;padding:4px 8px;font-size:11px;color:var(--text1);font-family:'DM Sans',sans-serif;outline:none;cursor:pointer;}
.chart-opt-select:focus{border-color:var(--accent);}
.chart-trend-multi{min-width:150px;max-width:190px;height:auto;vertical-align:top;}
.chart-trend-multi option{padding:1px 4px;}
.chart-trend-multi option:checked{background:var(--accent);color:#fff;}
.chart-opt-toggle{padding:4px 10px;border-radius:5px;font-size:11px;font-weight:500;cursor:pointer;border:0.5px solid var(--nav-border);background:var(--toggle-bg);color:var(--text3);font-family:'DM Sans',sans-serif;transition:all 0.15s;white-space:nowrap;}
.chart-opt-toggle:hover{color:var(--text2);}
.chart-opt-toggle.active{background:var(--tab-bg);color:var(--accent-text);border-color:var(--accent);}
.chart-divider{width:0.5px;height:18px;background:var(--nav-border);margin:0 4px;}
/* ── COLUMN FILTERS ── */
.col-filter-icon{display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;border-radius:3px;cursor:pointer;margin-left:4px;font-size:9px;color:var(--text3);transition:all 0.15s;vertical-align:middle;}
.col-filter-icon:hover{color:var(--accent);}
.col-filter-icon.active{color:var(--accent);background:var(--tab-bg);}
.filter-dropdown{position:fixed;z-index:9999;background:var(--surface);border:1px solid var(--nav-border);border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,0.35);min-width:200px;max-width:280px;overflow:hidden;}
.filter-dropdown-header{padding:8px 12px;border-bottom:0.5px solid var(--nav-border);font-size:11px;font-weight:600;color:var(--text2);display:flex;align-items:center;justify-content:space-between;}
.filter-dropdown-body{max-height:220px;overflow-y:auto;padding:6px 0;}
.filter-dropdown-footer{padding:6px 10px;border-top:0.5px solid var(--nav-border);display:flex;gap:6px;justify-content:flex-end;}
.filter-cat-item{display:flex;align-items:center;gap:7px;padding:4px 12px;cursor:pointer;font-size:11px;color:var(--text2);transition:background 0.1s;}
.filter-cat-item:hover{background:var(--toggle-bg);}
.filter-cat-item input{cursor:pointer;accent-color:var(--accent);}
.filter-num-row{padding:6px 12px;display:flex;align-items:center;gap:6px;}
.filter-num-select{background:var(--input-bg);border:0.5px solid var(--nav-border);border-radius:4px;padding:3px 6px;font-size:11px;color:var(--text1);font-family:'DM Sans',sans-serif;outline:none;}
.filter-num-input{background:var(--input-bg);border:0.5px solid var(--nav-border);border-radius:4px;padding:3px 8px;font-size:11px;color:var(--text1);font-family:'DM Sans',sans-serif;outline:none;width:90px;}
.filter-num-input:focus{border-color:var(--accent);}
.filter-action-btn{padding:3px 10px;border-radius:4px;font-size:10px;font-weight:600;cursor:pointer;border:0.5px solid var(--nav-border);font-family:'DM Sans',sans-serif;transition:all 0.15s;}
.filter-apply-btn{background:var(--accent);color:var(--accent-text);border-color:var(--accent);}
.filter-apply-btn:hover{opacity:0.85;}
.filter-clear-btn{background:var(--toggle-bg);color:var(--text2);}
.filter-clear-btn:hover{color:var(--text1);}
.filter-search-input{margin:6px 10px;padding:4px 8px;border-radius:4px;border:0.5px solid var(--nav-border);background:var(--input-bg);color:var(--text1);font-size:11px;font-family:'DM Sans',sans-serif;outline:none;width:calc(100% - 20px);box-sizing:border-box;}
.filter-search-input:focus{border-color:var(--accent);}

/* ── CHART FILTER ROW ── */
.chart-filter-row{display:flex;gap:10px;margin-bottom:12px;align-items:flex-start;flex-wrap:wrap;}
.chart-filter-group{display:flex;flex-direction:column;flex:1;min-width:140px;position:relative;}
.chart-filter-tag{font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:3px;}
.chart-filter-trigger{background:var(--input-bg);border:0.5px solid var(--nav-border);border-radius:5px;padding:5px 9px;font-size:11px;color:var(--text1);font-family:'DM Sans',sans-serif;cursor:pointer;display:flex;align-items:center;justify-content:space-between;gap:6px;transition:border-color 0.15s;user-select:none;}
.chart-filter-trigger:hover{border-color:var(--accent);}
.chart-filter-trigger.active{border-color:var(--accent);color:var(--accent);}
.chart-filter-trigger-text{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.chart-filter-dd{position:fixed;z-index:9999;background:var(--surface);border:1px solid var(--nav-border);border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,0.35);width:220px;overflow:hidden;}
.chart-filter-dd-search{margin:6px 8px;padding:4px 8px;border-radius:4px;border:0.5px solid var(--nav-border);background:var(--input-bg);color:var(--text1);font-size:11px;font-family:'DM Sans',sans-serif;outline:none;width:calc(100% - 16px);box-sizing:border-box;}
.chart-filter-dd-search:focus{border-color:var(--accent);}
.chart-filter-dd-body{max-height:180px;overflow-y:auto;padding:4px 0;}
.chart-filter-dd-item{display:flex;align-items:center;gap:7px;padding:4px 10px;cursor:pointer;font-size:11px;color:var(--text2);transition:background 0.1s;}
.chart-filter-dd-item:hover{background:var(--toggle-bg);}
.chart-filter-dd-item input{cursor:pointer;accent-color:var(--accent);flex-shrink:0;}
.chart-filter-dd-footer{padding:6px 8px;border-top:0.5px solid var(--nav-border);display:flex;gap:6px;justify-content:space-between;align-items:center;}
.chart-filter-count{font-size:10px;color:var(--text3);}
.chart-filter-apply{padding:3px 10px;border-radius:4px;font-size:10px;font-weight:600;cursor:pointer;background:var(--accent);color:var(--accent-text);border:none;font-family:'DM Sans',sans-serif;}
.chart-filter-clear{padding:3px 10px;border-radius:4px;font-size:10px;font-weight:600;cursor:pointer;background:var(--toggle-bg);color:var(--text2);border:0.5px solid var(--nav-border);font-family:'DM Sans',sans-serif;}
.cft-multi-row{display:flex;align-items:center;gap:8px;margin-bottom:8px;flex-wrap:wrap;}
.cft-multi-row-num{font-size:10px;color:var(--text3);font-weight:600;min-width:14px;}
.cft-remove-btn{padding:2px 7px;border-radius:4px;font-size:11px;cursor:pointer;background:transparent;border:0.5px solid var(--nav-border);color:var(--text3);font-family:'DM Sans',sans-serif;flex-shrink:0;transition:all 0.15s;}
.cft-remove-btn:hover{color:#f75a7a;border-color:#f75a7a;}
.cft-multi-group{display:flex;flex-direction:column;flex:1;min-width:120px;}
.cft-hdr{display:flex;align-items:center;gap:8px;margin-bottom:8px;}
.cft-eye{background:transparent;border:none;cursor:pointer;font-size:13px;padding:0 2px;line-height:1;filter:grayscale(0);}
.cft-multi-row.cft-off{opacity:.45;}
.cft-multi-row.cft-off .chart-filter-trigger{text-decoration:line-through;}
.cft-badge{font-size:8px;font-weight:700;background:var(--surface2);border:0.5px solid var(--border);border-radius:3px;padding:0 3px;color:var(--text3);margin-right:5px;}
.cft-range-row{display:flex;align-items:center;gap:8px;padding:5px 10px;}
.cft-range-row label{font-size:11px;color:var(--text2);width:30px;}
.cft-range-row input{flex:1;background:var(--input-bg);border:0.5px solid var(--nav-border);border-radius:4px;padding:4px 8px;font-size:11px;color:var(--text1);outline:none;}
.cft-range-row input:focus{border-color:var(--accent);}

/* ── CHART THEME ROW ── */
.chart-theme-row{display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap;padding:8px 10px;background:var(--toggle-bg);border-radius:6px;border:0.5px solid var(--nav-border);}
.chart-theme-label{font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;margin-right:2px;}
.palette-btn{width:18px;height:18px;border-radius:4px;cursor:pointer;border:2px solid transparent;transition:all 0.15s;position:relative;overflow:hidden;flex-shrink:0;}
.palette-btn.active{border-color:var(--text1);transform:scale(1.15);}
.palette-btn:hover{transform:scale(1.1);}

/* ── SIDEBAR CHART LINKS ── */
.sidebar-chart-link{display:flex;align-items:center;gap:6px;padding:3px 12px 3px 28px;font-size:11px;color:var(--text3);cursor:pointer;transition:all 0.15s;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.sidebar-chart-link:hover{color:var(--accent);}
.sidebar-chart-link::before{content:'📊';font-size:10px;flex-shrink:0;}

/* ── FLEXIBLE COMBO PANEL ── */
.combo-axes-panel{display:none;gap:8px;margin-bottom:10px;}
.combo-axes-panel.visible{display:flex;}
.combo-axis-half{flex:1;background:var(--toggle-bg);border:0.5px solid var(--nav-border);border-radius:6px;padding:8px 10px;min-width:0;}
.combo-axis-half-label{font-size:10px;font-weight:700;color:var(--text2);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;display:flex;align-items:center;gap:6px;}
.combo-axis-badge{font-size:9px;padding:2px 6px;border-radius:3px;font-weight:600;}
.combo-axis-badge.left{background:#4f8ef720;color:#4f8ef7;}
.combo-axis-badge.right{background:#f7b24f20;color:#f7b24f;}
.combo-type-group{display:flex;gap:3px;margin-bottom:6px;}
.combo-type-btn{padding:3px 8px;border-radius:4px;font-size:10px;font-weight:500;cursor:pointer;border:0.5px solid var(--nav-border);background:var(--toggle-bg);color:var(--text3);font-family:'DM Sans',sans-serif;transition:all 0.15s;}
.combo-type-btn:hover{color:var(--text2);}
.combo-type-btn.active{background:var(--tab-bg);color:var(--accent-text);border-color:var(--accent);}
.combo-y-checks{display:flex;flex-direction:column;gap:3px;max-height:120px;overflow-y:auto;}
</style>
</head>
<body>
<div class="app">
  <aside class="sidebar">
    <h2>📊 pudbo · Polars</h2>
    <div class="proj" style="display:flex;gap:4px;align-items:center;padding:0 10px 8px;">
      <select id="projSel" onchange="openProject(this.value)" title="Switch project" style="flex:1;min-width:0;"></select>
      <button class="btn" onclick="newProject()" title="New project">＋</button>
      <button class="btn" onclick="renameProject()" title="Rename current project">✎</button>
      <button class="btn" onclick="deleteProject()" title="Delete current project">🗑</button>
    </div>
    <div class="up">
      <label>⬆ Upload CSV / Parquet / JSON
        <input type="file" id="file" accept=".csv,.tsv,.parquet,.pq,.json,.ndjson,.jsonl" multiple/>
      </label>
    </div>
    <div class="tlist" id="tlist"><div class="muted" style="padding:10px;">No tables yet — upload a file.</div></div>
  </aside>
  <main class="main">
    <div class="toolbar">
      <button class="btn sb-toggle" onclick="toggleSidebar()" title="Show / hide sidebar">☰</button>
      <input id="curName" placeholder="—" title="Rename table — type and press Enter"
             onkeydown="if(event.key==='Enter')this.blur()" onchange="renameTable(this.value)"/>
      <span class="spinner" id="spinner" title="Working…"></span>
      <span style="margin-right:auto;"></span>
      <span class="muted">Theme</span>
      <select id="themeSel" onchange="setTheme(this.value)" title="Colour theme"></select>
      <span class="muted">Numbers</span>
      <div class="seg" id="fmtseg">
        <button class="on" data-f="actual">Actual</button>
        <button data-f="k">K</button><button data-f="m">M</button><button data-f="b">B</button>
      </div>
      <button class="btn" onclick="openClean()">🧹 Clean</button>
      <button class="btn" onclick="openFeature()">🛠 Feature</button>
      <button class="btn" onclick="openProfile()">🔎 Profile</button>
      <button class="btn" onclick="openGroupby()">⊕ Group By</button>
      <button class="btn" onclick="openPivot()">⊞ Pivot</button>
      <button class="btn" onclick="openJoin()">⇄ Relationship</button>
      <button class="btn" onclick="openUnion()">⊔ Union</button>
      <button class="btn good" onclick="addNewChart()">📈 Add Chart</button>
      <button class="btn" id="recipeBtn" onclick="openRecipe()" title="Replay clean/feature steps on another table">📋 Recipe</button>
      <button class="btn" id="undoBtn" onclick="undoTx()" title="Undo last transform">↶ Undo</button>
      <select id="exportSel" title="Export current table" onchange="exportAs(this.value);this.selectedIndex=0;">
        <option>↓ Export</option><option value="csv">CSV</option><option value="parquet">Parquet</option><option value="json">JSON</option><option value="xlsx">Excel</option>
      </select>
    </div>
    <div class="scroll" id="scroll">
      <div class="content" id="content"><div class="empty">⬅ Upload data to begin.<div class="muted">All transforms run server-side in Polars.</div></div></div>
      <div class="chartpanel" id="chartpanel" style="display:none;">
        <div class="chartpanel-hdr"><span id="chartpanel-title">Charts</span></div>
        <div id="chart-instances"></div>
      </div>
    </div>
  </main>
</div>

<div class="overlay" id="overlay"><div class="modal" id="modal"></div></div>
<div class="toast" id="toast"></div>

<script>
const API = '';
let TABLES = [], CUR = null, NUMFMT = 'actual', _GRIDCOLS = [];
const state = {}; // per-table: {page,sort:[],search,filters:[]}

function st(n){ if(!state[n]) state[n]={page:1,sort:[],search:'',filters:[],page_size:50}; return state[n]; }
function toast(msg,bad){ const t=document.getElementById('toast'); t.textContent=msg; t.className='toast'+(bad?' bad':''); t.style.display='block'; setTimeout(()=>t.style.display='none',2600); }
function esc(s){ return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
// ── global loading indicator (shown whenever a backend call is in flight) ──
let _busy=0;
function busy(on){ _busy=Math.max(0,_busy+(on?1:-1)); const s=document.getElementById('spinner'); if(s) s.style.display=_busy>0?'inline-block':'none'; }
async function jget(u){ busy(1); try{ const r=await fetch(API+u); if(!r.ok) throw new Error((await r.json()).detail||r.statusText); return await r.json(); } finally{ busy(0); } }
async function jpost(u,b){ busy(1); try{ const r=await fetch(API+u,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)}); if(!r.ok) throw new Error((await r.json()).detail||r.statusText); return await r.json(); } finally{ busy(0); } }

// ── themes ──
const THEMES=[['','Dark Blue'],['theme-carbon','Carbon'],['theme-midnight','Midnight'],['theme-crimson','Crimson'],['theme-grape','Grape'],['theme-slate','Slate Light']];
function setTheme(cls){ document.body.className=cls||''; try{ localStorage.setItem('pudbo_theme',cls||''); }catch(e){} }
function initTheme(){
  const sel=document.getElementById('themeSel');
  sel.innerHTML=THEMES.map(t=>`<option value="${t[0]}">${t[1]}</option>`).join('');
  let saved=''; try{ saved=localStorage.getItem('pudbo_theme')||''; }catch(e){}
  sel.value=saved; setTheme(saved);
  let sb=''; try{ sb=localStorage.getItem('pudbo_sb')||''; }catch(e){}
  if(sb==='1') document.querySelector('.app').classList.add('sidebar-collapsed');
}

// ── collapsible sidebar ──
function toggleSidebar(){
  const c=document.querySelector('.app').classList.toggle('sidebar-collapsed');
  try{ localStorage.setItem('pudbo_sb', c?'1':'0'); }catch(e){}
}

// ── rename the current table ──
async function renameTable(newName){
  if(!CUR) return;
  newName=(newName||'').trim();
  if(!newName || newName===CUR){ document.getElementById('curName').value=CUR||''; return; }
  try{
    const r=await jpost('/api/table/rename',{old:CUR,new:newName});
    CUR=r.name; toast('Renamed → '+r.name); await refreshTables(r.name);
  }catch(err){ toast(err.message,true); document.getElementById('curName').value=CUR||''; }
}

// ── number formatting for the DATA GRID (display only; Polars holds real values) ──
function fmtNum(v, numeric){
  if(v===null||v===undefined) return '—';
  if(!numeric) return esc(v);
  const n = Number(v); if(!isFinite(n)) return esc(v);
  if(NUMFMT==='actual') return n.toLocaleString('en-IN');
  const a=Math.abs(n);
  if(NUMFMT==='k') return a<1e3?n.toLocaleString('en-IN'):(n/1e3).toFixed(2)+'K';
  if(NUMFMT==='m') return a<1e6?(a<1e3?n.toLocaleString('en-IN'):(n/1e3).toFixed(2)+'K'):(n/1e6).toFixed(2)+'M';
  if(NUMFMT==='b') return a<1e3?n.toLocaleString('en-IN'):a<1e6?(n/1e3).toFixed(2)+'K':a<1e9?(n/1e6).toFixed(2)+'M':(n/1e9).toFixed(2)+'B';
  return n.toLocaleString('en-IN');
}

// ── upload ──
document.getElementById('file').addEventListener('change', async e=>{
  for(const f of e.target.files){
    const fd=new FormData(); fd.append('file',f);
    busy(1);
    try{ const r=await fetch(API+'/api/upload',{method:'POST',body:fd});
      if(!r.ok){ toast('Upload failed: '+((await r.json()).detail||''),true); continue; }
      const j=await r.json(); toast('Loaded '+j.name+' ('+j.rows+'×'+j.cols+')');
    }catch(err){ toast('Upload error: '+err.message,true); }
    finally{ busy(0); }
  }
  e.target.value=''; await refreshTables();
});

async function refreshTables(selectName){
  TABLES = await jget('/api/tables');
  const el=document.getElementById('tlist');
  if(!TABLES.length){ el.innerHTML='<div class="muted" style="padding:10px;">No tables yet — upload a file.</div>'; return; }
  el.innerHTML = TABLES.map(t=>{
    const nm=JSON.stringify(t.name);
    const links=(t.charts||[]).map(c=>`<div class="chart-nav-link" onclick='event.stopPropagation();openSavedChart(${nm},${JSON.stringify(c.id)})'>
        <span class="cnl-t">${esc(c.title)}</span>
        <span class="cnl-x" onclick='event.stopPropagation();deleteSavedChart(${nm},${JSON.stringify(c.id)},event)'>✕</span>
      </div>`).join('');
    return `<div class="twrap">
      <div class="titem ${t.name===CUR?'active':''}" onclick='selectTable(${nm})'>
        <div class="kind">${esc(t.kind)}</div>
        <div class="nm">${esc(t.name)}</div>
        <div class="mt">${t.rows.toLocaleString()} rows · ${t.cols} cols ${t.detail?('· '+esc(t.detail)):''}
          <span style="float:right;cursor:pointer;color:var(--bad);" onclick='event.stopPropagation();dropTable(${nm})'>✕</span></div>
      </div>
      ${links?`<div class="charts-nav">${links}</div>`:''}
    </div>`;
  }).join('');
  if(selectName) selectTable(selectName);
  else if(CUR && TABLES.some(t=>t.name===CUR)) renderTable();
  else if(TABLES.length && !CUR) selectTable(TABLES[0].name);
}

async function dropTable(n){ busy(1); try{ await fetch(API+'/api/table/'+encodeURIComponent(n),{method:'DELETE'}); } finally{ busy(0); } if(CUR===n){CUR=null;document.getElementById('content').innerHTML='';document.getElementById('curName').value='';document.getElementById('chartpanel').style.display='none';document.getElementById('chart-instances').innerHTML='';} refreshTables(); }

// ── projects (VSCode-style workspaces; each has its own tables/charts) ──
let PROJECTS_LIST = [], ACTIVE_PID = null;
async function refreshProjects(){
  let d; try{ d=await jget('/api/projects'); }catch(e){ return; }
  PROJECTS_LIST=d.projects||[]; ACTIVE_PID=d.active;
  const sel=document.getElementById('projSel');
  if(sel) sel.innerHTML=PROJECTS_LIST.map(p=>`<option value="${esc(p.id)}" ${p.id===ACTIVE_PID?'selected':''}>${esc(p.name)} · ${p.tables} tbl</option>`).join('');
}
function _resetWorkspaceUI(){
  CUR=null;
  for(const k in state) delete state[k];
  document.getElementById('content').innerHTML='<div class="empty">⬅ Upload data to begin.<div class="muted">All transforms run server-side in Polars.</div></div>';
  document.getElementById('curName').value='';
  document.getElementById('chartpanel').style.display='none';
  document.getElementById('chart-instances').innerHTML='';
  _CHART_DATA=null; _CHART_TABLE=null;
}
async function openProject(id){
  if(!id || id===ACTIVE_PID) return;
  try{ await jpost('/api/projects/open',{id}); }catch(e){ toast(e.message,true); return; }
  _resetWorkspaceUI(); await refreshProjects(); await refreshTables(); toast('Opened project');
}
async function newProject(){
  const name=prompt('New project name:'); if(name===null) return;
  try{ const r=await jpost('/api/projects/create',{name:name||'Untitled'}); ACTIVE_PID=r.active; }catch(e){ toast(e.message,true); return; }
  _resetWorkspaceUI(); await refreshProjects(); await refreshTables(); toast('Created project');
}
async function renameProject(){
  const cur=PROJECTS_LIST.find(p=>p.id===ACTIVE_PID);
  const name=prompt('Rename project:', cur?cur.name:''); if(name===null||!name.trim()) return;
  try{ await jpost('/api/projects/rename',{id:ACTIVE_PID,name:name.trim()}); }catch(e){ toast(e.message,true); return; }
  await refreshProjects(); toast('Renamed');
}
async function deleteProject(){
  const cur=PROJECTS_LIST.find(p=>p.id===ACTIVE_PID);
  if(!confirm('Delete project "'+(cur?cur.name:ACTIVE_PID)+'" and ALL its tables/charts? This cannot be undone.')) return;
  try{ await jpost('/api/projects/delete',{id:ACTIVE_PID}); }catch(e){ toast(e.message,true); return; }
  _resetWorkspaceUI(); await refreshProjects(); await refreshTables(); toast('Deleted');
}
function tableMeta(n){ return TABLES.find(t=>t.name===n); }
function selectTable(n){ CUR=n; document.getElementById('curName').value=n; document.querySelectorAll('.titem').forEach(e=>e.classList.remove('active')); st(n).page=1; renderTable(); refreshActive(); renderChartPanel(); }
function refreshActive(){ document.querySelectorAll('.titem').forEach(e=>{ e.classList.toggle('active', e.querySelector('.nm')?.textContent===CUR); }); }
function updateUndoBtn(){
  const t=tableMeta(CUR);
  const b=document.getElementById('undoBtn');
  if(b){ const n=t?(t.history||0):0; b.textContent='↶ Undo'+(n?(' ('+n+')'):''); b.style.opacity=n?1:.45; b.title=(t&&t.undo)?('Undo: '+t.undo):'Nothing to undo'; }
  const rb=document.getElementById('recipeBtn');
  if(rb){ const rn=t?(t.recipe||0):0; rb.textContent='📋 Recipe'+(rn?(' ('+rn+')'):''); rb.style.opacity=rn?1:.55; rb.title=rn?('Replay '+rn+' recorded step(s) onto another table'):'No recorded clean/feature steps yet'; }
}

// ── data grid (Polars does filter/sort/paginate/summary) ──
async function renderTable(){
  if(!CUR) return;
  const s=st(CUR);
  updateUndoBtn();
  let d;
  try{ d=await jpost('/api/data',{table:CUR,page:s.page,page_size:s.page_size,sort:s.sort,search:s.search,filters:s.filters}); }
  catch(err){ toast(err.message,true); return; }
  const numericSet=new Set(d.numeric_cols);
  const sortDir=c=>{ const k=s.sort.find(x=>x.col===c); return k?(k.dir>0?' ▲':' ▼'):''; };
  _GRIDCOLS=d.cols;
  const filteredCols=new Set(s.filters.map(f=>f.col));
  const head='<tr>'+d.cols.map((c,i)=>`<th class="${numericSet.has(c)?'num':''}"><span class="th-lbl" onclick="cycleSort(_GRIDCOLS[${i}])">${esc(c)}<span class="ar">${sortDir(c)}</span></span><span class="th-menu${filteredCols.has(c)?' on':''}" onclick="openColMenu(${i},event)" title="Filter / sort this column">▾</span></th>`).join('')+'</tr>';
  const body=d.rows.map(r=>'<tr>'+r.map((v,i)=>{const num=numericSet.has(d.cols[i]);return `<td class="${num?'num':''}">${fmtNum(v,num)}</td>`;}).join('')+'</tr>').join('');
  const sumRow='<tr class="sum">'+d.summary.map((v,i)=>{const num=numericSet.has(d.cols[i]);return `<td class="${num?'num':''}">${i===0&&v===null?'∑ Total':fmtNum(v,num)}</td>`;}).join('')+'</tr>';
  const filtChips=s.filters.map((f,i)=>{
    let lbl;
    if(f.op==='in') lbl=esc(f.col)+': '+(Array.isArray(f.value)?f.value.length:0)+' selected';
    else if(f.op==='range'){ const v=f.value||{}; lbl=esc(f.col)+' '+((v.min!=null?'≥'+v.min:'')+(v.min!=null&&v.max!=null?' ':'')+(v.max!=null?'≤'+v.max:'')); }
    else if(f.op==='between'&&Array.isArray(f.value)) lbl=esc(f.col)+' '+esc(f.value[0])+'…'+esc(f.value[1]);
    else lbl=esc(f.col)+' '+esc(f.op)+' '+esc(f.value);
    return `<span class="chip" title="${esc(f.col)} filter">${lbl} <span style="cursor:pointer" onclick="rmFilter(${i})">✕</span></span>`;
  }).join(' ');
  document.getElementById('content').innerHTML=`
    <div style="display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap;">
      <input placeholder="🔍 Search all columns…" value="${esc(s.search)}" oninput="onSearch(this.value)" style="flex:1;min-width:200px;"/>
      <button class="btn" onclick="openFilter()">+ Filter</button>
      <span class="muted">${d.total.toLocaleString()} rows</span>
    </div>
    <div>${filtChips}</div>
    <div style="overflow:auto;max-height:calc(100vh - 230px);border:1px solid var(--border);border-radius:8px;margin-top:8px;">
      <table>${head}${body}${d.total?sumRow:''}</table>
    </div>
    <div class="pager">
      <button class="btn" ${d.page<=1?'disabled':''} onclick="goPage(${d.page-1})">‹ Prev</button>
      <span class="muted">Page ${d.page} / ${d.pages}</span>
      <button class="btn" ${d.page>=d.pages?'disabled':''} onclick="goPage(${d.page+1})">Next ›</button>
      <select onchange="setRpp(this.value)" style="margin-left:8px;">${[25,50,100,250].map(n=>`<option ${n===s.page_size?'selected':''}>${n}</option>`).join('')}</select>
    </div>`;
}
let _searchT;
function onSearch(v){ const s=st(CUR); s.search=v; s.page=1; clearTimeout(_searchT); _searchT=setTimeout(()=>{ renderTable(); refreshChartsForFilter(); },250); }
function cycleSort(c){ const s=st(CUR); const k=s.sort.find(x=>x.col===c); if(!k) s.sort=[{col:c,dir:1}]; else if(k.dir>0) s.sort=[{col:c,dir:-1}]; else s.sort=[]; renderTable(); }
function goPage(p){ const s=st(CUR); s.page=p; renderTable(); }
function setRpp(v){ const s=st(CUR); s.page_size=+v; s.page=1; renderTable(); }
function rmFilter(i){ const s=st(CUR); s.filters.splice(i,1); s.page=1; renderTable(); refreshChartsForFilter(); }
document.getElementById('fmtseg').addEventListener('click',e=>{ if(e.target.dataset.f){ NUMFMT=e.target.dataset.f; document.querySelectorAll('#fmtseg button').forEach(b=>b.classList.toggle('on',b===e.target)); renderTable(); } });
function exportAs(fmt){ if(!fmt||!CUR) return; window.location=API+'/api/export/'+encodeURIComponent(CUR)+'?fmt='+encodeURIComponent(fmt); }

// ── column header filter dropdown (db_fw12-style: value checklist / numeric range + actions) ──
let _colMenuEl=null;
function _closeColMenu(){ if(_colMenuEl){ _colMenuEl.remove(); _colMenuEl=null; } }
document.addEventListener('click',e=>{ if(_colMenuEl && !_colMenuEl.contains(e.target) && !e.target.classList.contains('th-menu')) _closeColMenu(); });

async function openColMenu(i, ev){
  ev.stopPropagation(); _closeColMenu();
  const col=_GRIDCOLS[i]; if(col===undefined||!CUR) return;
  let data;
  try{ data=await jpost('/api/values',{table:CUR,col,limit:2000}); }
  catch(err){ toast(err.message,true); return; }
  const existing=st(CUR).filters.find(f=>f.col===col);
  const dd=document.createElement('div'); dd.className='colfilter'; dd.dataset.col=col; dd.dataset.idx=i;
  let html=`<div class="cf-head"><span class="cf-title">${esc(col)}</span>
      <span class="cf-sub">${data.numeric?('range '+fmtMaybe(data.min)+'…'+fmtMaybe(data.max)):(data.total_unique.toLocaleString()+' unique')}</span></div>
    <div class="cf-actions">
      <button class="cf-act" onclick="colSortFromMenu(${i},1)" title="Sort ascending">↑</button>
      <button class="cf-act" onclick="colSortFromMenu(${i},-1)" title="Sort descending">↓</button>
      <button class="cf-act" onclick="colSortFromMenu(${i},0)" title="Clear sort">⨯sort</button>
      <span style="flex:1"></span>
      <button class="cf-act" onclick="_closeColMenu();colRename(${i})" title="Rename column">✎</button>
      <button class="cf-act" onclick="_closeColMenu();colCast(${i},'float')" title="Cast to number">#</button>
      <button class="cf-act" onclick="_closeColMenu();colCast(${i},'date')" title="Cast to date">📅</button>
      <button class="cf-act" onclick="_closeColMenu();colDrop(${i})" title="Drop column">🗑</button>
    </div>`;
  if(data.numeric){
    const r=(existing&&existing.op==='range')?existing.value:{};
    html+=`<div class="cf-range"><label>Min</label><input type="number" step="any" id="cf-min" value="${r.min!=null?r.min:''}" placeholder="${fmtMaybe(data.min)}">
      <label>Max</label><input type="number" step="any" id="cf-max" value="${r.max!=null?r.max:''}" placeholder="${fmtMaybe(data.max)}"></div>`;
  }else{
    const allowed=(existing&&existing.op==='in')?new Set(existing.value.map(String)):null;  // null = all
    html+=`<input class="cf-search" placeholder="Search values…" oninput="cfSearch(this)">
      <div class="cf-body" id="cf-body">
        <label class="cf-item cf-all"><input type="checkbox" id="cf-selall" ${allowed===null?'checked':''} onchange="cfToggleAll(this)"> <b>Select All</b></label>
        ${data.values.map(v=>`<label class="cf-item" data-v="${esc(v)}"><input type="checkbox" value="${esc(v)}" ${(allowed===null||allowed.has(v))?'checked':''} onchange="cfItemCheck()"> ${esc(v)}</label>`).join('')}
      </div>
      ${data.shown<data.total_unique?`<div class="cf-note">first ${data.shown.toLocaleString()} of ${data.total_unique.toLocaleString()} — use search</div>`:''}`;
  }
  html+=`<div class="cf-foot"><button class="cf-clear" onclick="cfClear()">Clear</button><button class="cf-apply" onclick="cfApply()">Apply</button></div>`;
  dd.innerHTML=html;
  document.body.appendChild(dd);
  dd.style.left=Math.min(ev.clientX, window.innerWidth-dd.offsetWidth-8)+'px';
  dd.style.top=Math.min(ev.clientY, window.innerHeight-dd.offsetHeight-8)+'px';
  _colMenuEl=dd;
  const se=dd.querySelector('.cf-search'); if(se) se.focus();
}
function cfSearch(inp){ const q=inp.value.toLowerCase(); inp.closest('.colfilter').querySelectorAll('.cf-item[data-v]').forEach(el=>{ el.style.display=el.dataset.v.toLowerCase().includes(q)?'':'none'; }); }
function cfToggleAll(cb){ const dd=cb.closest('.colfilter'); dd.querySelectorAll('.cf-body input[value]').forEach(c=>{ if(c.closest('.cf-item').style.display!=='none') c.checked=cb.checked; }); }
function cfItemCheck(){ const dd=_colMenuEl; if(!dd) return; const all=dd.querySelector('#cf-selall'); const items=[...dd.querySelectorAll('.cf-body input[value]')]; if(all) all.checked=items.every(c=>c.checked); }
function cfApply(){
  const dd=_colMenuEl; if(!dd) return; const col=dd.dataset.col, s=st(CUR);
  s.filters=s.filters.filter(f=>f.col!==col);   // one header-filter per column (replace)
  const minEl=dd.querySelector('#cf-min');
  if(minEl){
    const mn=minEl.value, mx=dd.querySelector('#cf-max').value;
    if(mn!==''||mx!=='') s.filters.push({col,op:'range',value:{min:mn===''?null:parseFloat(mn),max:mx===''?null:parseFloat(mx)}});
  }else{
    const selall=dd.querySelector('#cf-selall');
    if(!(selall&&selall.checked)){   // Select-All checked = keep everything (no filter)
      const checked=[...dd.querySelectorAll('.cf-body input[value]:checked')].map(c=>c.value);
      s.filters.push({col,op:'in',value:checked});   // [] = show none
    }
  }
  s.page=1; _closeColMenu(); renderTable(); refreshChartsForFilter();
}
function cfClear(){
  const dd=_colMenuEl; if(!dd) return; const col=dd.dataset.col, s=st(CUR);
  s.filters=s.filters.filter(f=>f.col!==col); s.page=1;
  _closeColMenu(); renderTable(); refreshChartsForFilter();
}
function colSortFromMenu(i,dir){ _closeColMenu(); setColSort(i,dir); }
function setColSort(i,dir){ const c=_GRIDCOLS[i],s=st(CUR); s.sort= dir===0?[]:[{col:c,dir}]; renderTable(); refreshChartsForFilter(); }
function colRename(i){ const c=_GRIDCOLS[i]; const to=prompt('Rename column "'+c+'" to:', c); if(to&&to.trim()&&to!==c) runTx('rename',{col:c,to:to.trim()}); }
function colCast(i,dtype){ runTx('cast',{col:_GRIDCOLS[i],dtype}); }
function colDrop(i){ const c=_GRIDCOLS[i]; if(confirm('Drop column "'+c+'" from '+CUR+'?')) runTx('dropcols',{cols:[c]}); }

// ── modal helpers ──
function showModal(html){ document.getElementById('modal').innerHTML=html+'<div style="text-align:right;margin-top:14px;"><button class="btn" onclick="closeModal()">Close</button></div>'; document.getElementById('overlay').classList.add('open'); }
function closeModal(){ document.getElementById('overlay').classList.remove('open'); document.getElementById('modal').style.width=''; }
document.getElementById('overlay').addEventListener('click',e=>{ if(e.target.id==='overlay') closeModal(); });
function colsOf(n){ return (tableMeta(n)?.schema||[]).map(c=>c.name); }
function numColsOf(n){ return (tableMeta(n)?.schema||[]).filter(c=>c.numeric).map(c=>c.name); }
function opts(arr,sel){ return arr.map(c=>`<option ${c===sel?'selected':''}>${esc(c)}</option>`).join(''); }

// ── Filter ──
function openFilter(){
  if(!CUR) return;
  showModal(`<h3>Add filter</h3>
    <div class="row">
      <div class="field"><label>Column</label><select id="f-col">${opts(colsOf(CUR))}</select></div>
      <div class="field"><label>Operator</label><select id="f-op">${['=','!=','>','>=','<','<=','contains','between'].map(o=>`<option>${o}</option>`).join('')}</select></div>
      <div class="field"><label>Value</label><input id="f-val" placeholder="value"/></div>
      <div class="field" id="f-val2wrap" style="display:none;"><label>… to</label><input id="f-val2"/></div>
    </div>
    <button class="btn accent" onclick="addFilter()">Apply filter</button>`);
  document.getElementById('f-op').addEventListener('change',e=>{ document.getElementById('f-val2wrap').style.display = e.target.value==='between'?'':'none'; });
}
function addFilter(){
  const col=document.getElementById('f-col').value, op=document.getElementById('f-op').value;
  let value=document.getElementById('f-val').value;
  if(op==='between') value=[value, document.getElementById('f-val2').value];
  st(CUR).filters.push({col,op,value}); st(CUR).page=1; closeModal(); renderTable(); refreshChartsForFilter();
}

// ── Group By (Polars group_by — MULTIPLE metrics) ──
function openGroupby(){
  if(!CUR) return;
  const cols=colsOf(CUR);
  showModal(`<h3>⊕ Group By <span class="muted" style="font-size:11px;">— Polars group_by/agg · multiple metrics</span></h3>
    <div class="field"><label>Group by (Ctrl/Cmd-click for multiple)</label><select id="g-by" multiple>${opts(cols)}</select></div>
    <div class="subhdr">Metrics — add as many as you like</div>
    <div id="g-metrics"></div>
    <button class="minibtn" onclick="gbAddMetric()">+ Add metric</button>
    <div style="margin-top:14px;"><button class="btn accent" onclick="runGroupby()">▶ Build grouped table</button></div>`);
  gbAddMetric();
}
function gbAddMetric(){
  const nums=numColsOf(CUR);
  const wrap=document.getElementById('g-metrics');
  const div=document.createElement('div'); div.className='row gm'; div.style.marginBottom='6px';
  div.innerHTML=`<div class="field" style="margin-bottom:0;"><select class="gm-col">${opts(nums)}</select></div>
    <div class="field" style="margin-bottom:0;max-width:150px;"><select class="gm-agg">${['sum','mean','min','max','count','median','std','first','last','n_unique'].map(a=>`<option>${a}</option>`).join('')}</select></div>
    <button class="minibtn" onclick="this.parentElement.remove()">✕</button>`;
  wrap.appendChild(div);
}
async function runGroupby(){
  const by=[...document.getElementById('g-by').selectedOptions].map(o=>o.value);
  if(!by.length){ toast('Pick at least one group column',true); return; }
  const aggs=[...document.querySelectorAll('#g-metrics .gm')].map(r=>({col:r.querySelector('.gm-col').value, agg:r.querySelector('.gm-agg').value})).filter(a=>a.col);
  try{ const r=await jpost('/api/groupby',{table:CUR,by,aggs}); closeModal(); toast('Grouped → '+r.name); await refreshTables(r.name); }
  catch(err){ toast(err.message,true); }
}

// ── Pivot (Polars pivot, multi index + multi columns) ──
function openPivot(){
  if(!CUR) return;
  const cols=colsOf(CUR), nums=numColsOf(CUR);
  showModal(`<h3>⊞ Pivot <span class="muted" style="font-size:11px;">— Polars pivot</span></h3>
    <div class="row">
      <div class="field"><label>Rows / index (multi)</label><select id="p-idx" multiple>${opts(cols)}</select></div>
      <div class="field"><label>Columns (multi, optional)</label><select id="p-col" multiple>${opts(cols)}</select></div>
    </div>
    <div class="row">
      <div class="field"><label>Values (multi)</label><select id="p-val" multiple>${opts(nums)}</select></div>
      <div class="field"><label>Aggregate</label><select id="p-agg">${['sum','mean','min','max','count','median'].map(a=>`<option>${a}</option>`).join('')}</select></div>
    </div>
    <button class="btn accent" onclick="runPivot()">▶ Build pivot</button>`);
}
async function runPivot(){
  const index=[...document.getElementById('p-idx').selectedOptions].map(o=>o.value);
  const columns=[...document.getElementById('p-col').selectedOptions].map(o=>o.value);
  const values=[...document.getElementById('p-val').selectedOptions].map(o=>o.value);
  const agg=document.getElementById('p-agg').value;
  if(!index.length||!values.length){ toast('Pick Rows and at least one Values column',true); return; }
  try{ const r=await jpost('/api/pivot',{table:CUR,index,columns,values,agg}); closeModal(); toast('Pivot → '+r.name); await refreshTables(r.name); }
  catch(err){ toast(err.message,true); }
}

// ── Union / append (stack tables) ──
function openUnion(){
  if(TABLES.length<2){ toast('Upload at least 2 tables to stack',true); return; }
  showModal(`<h3>⊔ Union <span class="muted" style="font-size:11px;">— stack tables on top of each other</span></h3>
    <div class="field"><label>Tables to stack (Ctrl/Cmd-click — order matters)</label>
      <select id="un-tabs" multiple style="min-height:140px;">${TABLES.map(t=>`<option ${t.name===CUR?'selected':''}>${esc(t.name)}</option>`).join('')}</select></div>
    <div class="field"><label>Mode</label><select id="un-how">
      <option value="diagonal">Diagonal — keep all columns, fill gaps with null (different schemas OK)</option>
      <option value="vertical">Vertical — columns must match</option>
    </select></div>
    <button class="btn accent" onclick="runUnion()">▶ Build stacked table</button>`);
}
async function runUnion(){
  const tables=[...document.getElementById('un-tabs').selectedOptions].map(o=>o.value);
  const how=document.getElementById('un-how').value;
  if(tables.length<2){ toast('Pick at least two tables',true); return; }
  try{ const r=await jpost('/api/union',{tables,how}); closeModal(); toast('Union → '+r.name); await refreshTables(r.name); }
  catch(err){ toast(err.message,true); }
}

// ── Recipe (replay clean/feature steps on another table) ──
async function openRecipe(){
  if(!CUR){ toast('Select a table first',true); return; }
  let rc;
  try{ rc=await jget('/api/recipe/'+encodeURIComponent(CUR)); }catch(err){ toast(err.message,true); return; }
  const steps=rc.steps||[];
  const others=TABLES.filter(t=>t.name!==CUR).map(t=>t.name);
  const stepsHtml = steps.length
    ? steps.map((s,i)=>`<div style="font-size:12px;padding:3px 0;border-bottom:1px solid var(--border);"><span class="muted">${i+1}.</span> ${esc(s.label)}</div>`).join('')
    : '<div class="muted">No clean/feature steps recorded on this table yet. Apply some from 🧹 Clean or 🛠 Feature first.</div>';
  showModal(`<h3>📋 Recipe — ${esc(CUR)} <span class="muted" style="font-size:11px;">${steps.length} step(s)</span></h3>
    <div style="max-height:40vh;overflow:auto;border:1px solid var(--border);border-radius:8px;padding:8px 12px;margin-bottom:12px;">${stepsHtml}</div>
    ${steps.length&&others.length?`<div class="row">
      <div class="field"><label>Apply these steps to</label><select id="rc-target">${opts(others)}</select></div>
    </div>
    <button class="btn accent" onclick="runRecipe()">▶ Replay on selected table</button>`
    :(steps.length?'<div class="muted">Upload/select another table to replay onto.</div>':'')}`);
}
async function runRecipe(){
  const target=document.getElementById('rc-target').value;
  try{ const r=await jpost('/api/recipe/apply',{source:CUR,target}); closeModal(); toast('Replayed '+r.applied+' step(s) → '+target); await refreshTables(target); }
  catch(err){ toast(err.message,true); }
}

// ── Relationship / Join (multi-column mapping + keep-columns selection) ──
function openJoin(){
  if(TABLES.length<2){ toast('Upload at least 2 tables to relate',true); return; }
  const names=TABLES.map(t=>t.name);
  showModal(`<h3>⇄ Relationship <span class="muted" style="font-size:11px;">— Polars join · map on multiple columns</span></h3>
    <div class="row">
      <div class="field"><label>Left table</label><select id="j-l" onchange="jOnTableChange()">${opts(names,CUR)}</select></div>
      <div class="field"><label>Right table</label><select id="j-r" onchange="jOnTableChange()">${opts(names)}</select></div>
    </div>
    <div class="subhdr">Map columns to join on (left = right) — add a row per key pair</div>
    <div id="j-maps"></div>
    <button class="minibtn" id="j-addmap" onclick="jAddMap()">+ Add mapping column</button>
    <div class="row" style="margin-top:12px;">
      <div class="field"><label>Join type</label><select id="j-how" onchange="jHowChange()">${['inner','left','right','outer','cross'].map(h=>`<option>${h}</option>`).join('')}</select></div>
      <div class="field"><label>Cardinality (validated)</label><select id="j-val">
        <option value="m:m">many-to-many (m:m)</option>
        <option value="1:m">one-to-many (1:m)</option>
        <option value="m:1">many-to-one (m:1)</option>
        <option value="1:1">one-to-one (1:1)</option>
      </select></div>
    </div>
    <div class="subhdr">Keep columns (Ctrl/Cmd-click) — deselect to drop; leave all selected to keep everything</div>
    <div class="row">
      <div class="field"><label>From left</label><select id="j-keepl" multiple></select></div>
      <div class="field"><label>From right</label><select id="j-keepr" multiple></select></div>
    </div>
    <div class="muted" style="font-size:11px;margin-bottom:8px;">Cardinality is enforced by Polars — the join errors if the data violates it (e.g. duplicate keys on a 1:1).</div>
    <button class="btn accent" onclick="runJoin()">▶ Create related table</button>`);
  const r=document.getElementById('j-r'); r.value = names.find(n=>n!==CUR)||names[0];
  jOnTableChange();
}
function jHowChange(){
  const cross=document.getElementById('j-how').value==='cross';
  document.getElementById('j-maps').style.opacity = cross?0.4:1;
  document.getElementById('j-maps').style.pointerEvents = cross?'none':'';
  document.getElementById('j-addmap').style.display = cross?'none':'';
}
function _fillKeep(id, cols){ document.getElementById(id).innerHTML = cols.map(c=>`<option value="${esc(c)}" selected>${esc(c)}</option>`).join(''); }
function jOnTableChange(){
  const l=document.getElementById('j-l').value, r=document.getElementById('j-r').value;
  _fillKeep('j-keepl', colsOf(l));
  _fillKeep('j-keepr', colsOf(r));
  document.getElementById('j-maps').innerHTML='';
  jAddMap();
  jHowChange();
}
function jAddMap(){
  const l=document.getElementById('j-l').value, r=document.getElementById('j-r').value;
  const div=document.createElement('div'); div.className='row jmap'; div.style.marginBottom='6px'; div.style.alignItems='center';
  div.innerHTML=`<div class="field" style="margin-bottom:0;"><select class="jm-l">${opts(colsOf(l))}</select></div>
    <span style="color:var(--accent);font-weight:700;">=</span>
    <div class="field" style="margin-bottom:0;"><select class="jm-r">${opts(colsOf(r))}</select></div>
    <button class="minibtn" onclick="this.parentElement.remove()">✕</button>`;
  document.getElementById('j-maps').appendChild(div);
}
async function runJoin(){
  const left=document.getElementById('j-l').value, right=document.getElementById('j-r').value;
  const how=document.getElementById('j-how').value, validate=document.getElementById('j-val').value;
  const left_on=[...document.querySelectorAll('#j-maps .jmap .jm-l')].map(s=>s.value);
  const right_on=[...document.querySelectorAll('#j-maps .jmap .jm-r')].map(s=>s.value);
  const keepl=[...document.getElementById('j-keepl').selectedOptions].map(o=>o.value);
  const keepr=[...document.getElementById('j-keepr').selectedOptions].map(o=>o.value);
  const body={left,right,how,validate,left_on,right_on};
  if(keepl.length && keepl.length < colsOf(left).length) body.keep_left=keepl;
  if(keepr.length && keepr.length < colsOf(right).length) body.keep_right=keepr;
  try{ const r=await jpost('/api/join',body); closeModal(); toast('Related → '+r.name); await refreshTables(r.name); }
  catch(err){ toast(err.message,true); }
}

// ── small form helpers ──
function msel(id){ const e=document.getElementById(id); return e?[...e.selectedOptions].map(o=>o.value):[]; }
function val(id){ const e=document.getElementById(id); return e?e.value:undefined; }
function chk(id){ const e=document.getElementById(id); return !!(e&&e.checked); }
function fmtMaybe(v){ if(v===null||v===undefined) return '—'; if(typeof v==='number') return Math.abs(v)>=1000?v.toLocaleString('en-IN',{maximumFractionDigits:2}):(Math.round(v*1000)/1000); return esc(v); }

// ── apply an in-place transform, then refresh grid + sidebar (undo state) ──
async function runTx(op, params){
  if(!CUR){ toast('Select a table first',true); return; }
  try{
    const r=await jpost('/api/transform',{table:CUR,op,params,page_size:st(CUR).page_size});
    closeModal();
    st(CUR).page=1;                       // schema may have changed
    toast('Applied: '+r.label);
    await refreshTables();                // updates sidebar history + re-renders grid
    await renderChartPanel();             // data changed → reload charts from fresh rows
  }catch(err){ toast(err.message,true); }
}
async function undoTx(){
  if(!CUR) return;
  const t=tableMeta(CUR);
  if(!t || !t.history){ toast('Nothing to undo'); return; }
  try{ const r=await jpost('/api/undo',{table:CUR}); st(CUR).page=1; toast('Undone: '+r.undone); await refreshTables(); await renderChartPanel(); }
  catch(err){ toast(err.message,true); }
}

// ── Clean ──
const CLEAN_OPS=[['fillna','Fill missing'],['dropna','Drop missing'],['dropdup','Drop duplicates'],['cast','Cast type'],['strclean','Clean text'],['replace','Find & replace'],['outlier','Outliers'],['rename','Rename column'],['dropcols','Drop columns']];
function openClean(){ if(!CUR){ toast('Select a table first',true); return; }
  showModal(`<h3>🧹 Clean data <span class="muted" style="font-size:11px;">— in-place · undoable</span></h3>
    <div class="field"><label>Operation</label><select id="cl-op" onchange="cleanFields()">${CLEAN_OPS.map(o=>`<option value="${o[0]}">${o[1]}</option>`).join('')}</select></div>
    <div id="cl-fields"></div>
    <div style="margin-top:12px;"><button class="btn accent" onclick="applyClean()">▶ Apply</button></div>`);
  cleanFields();
}
function cleanFields(){
  const op=val('cl-op'), cols=colsOf(CUR), nums=numColsOf(CUR); let h='';
  if(op==='fillna') h=`<div class="row"><div class="field"><label>Columns</label><select id="cl-cols" multiple>${opts(cols)}</select></div>
    <div class="field"><label>Method</label><select id="cl-method" onchange="document.getElementById('cl-valwrap').style.display=this.value==='constant'?'':'none'">${['mean','median','mode','ffill','bfill','constant'].map(m=>`<option>${m}</option>`).join('')}</select></div>
    <div class="field" id="cl-valwrap" style="display:none;"><label>Value</label><input id="cl-val"/></div></div>`;
  else if(op==='dropna') h=`<div class="row"><div class="field"><label>What</label><select id="cl-axis">${[['rows','Drop rows that contain nulls'],['cols','Drop all-null columns']].map(a=>`<option value="${a[0]}">${a[1]}</option>`).join('')}</select></div>
    <div class="field"><label>Limit to columns (optional)</label><select id="cl-cols" multiple>${opts(cols)}</select></div></div>`;
  else if(op==='dropdup') h=`<div class="field"><label>Match on columns (blank = whole row)</label><select id="cl-cols" multiple>${opts(cols)}</select></div>`;
  else if(op==='cast') h=`<div class="row"><div class="field"><label>Column</label><select id="cl-col">${opts(cols)}</select></div>
    <div class="field"><label>To type</label><select id="cl-dtype">${['int','float','str','bool','date','datetime'].map(d=>`<option>${d}</option>`).join('')}</select></div>
    <div class="field"><label>Date format (optional)</label><input id="cl-fmt" placeholder="%Y-%m-%d"/></div></div>`;
  else if(op==='strclean') h=`<div class="field"><label>Text columns (blank = all text cols)</label><select id="cl-cols" multiple>${opts(cols)}</select></div>
    <div class="field"><label>Operations</label><div>${['trim','lower','upper','title','collapse'].map(o=>`<label style="margin-right:12px;font-size:12px;"><input type="checkbox" class="cl-strop" value="${o}" ${o==='trim'?'checked':''}> ${o}</label>`).join('')}</div></div>`;
  else if(op==='replace') h=`<div class="row"><div class="field"><label>Column</label><select id="cl-col">${opts(cols)}</select></div>
    <div class="field"><label>Find</label><input id="cl-find"/></div><div class="field"><label>Replace with</label><input id="cl-repl"/></div></div>
    <label style="font-size:12px;"><input type="checkbox" id="cl-regex"> regex</label>`;
  else if(op==='outlier') h=`<div class="row"><div class="field"><label>Column</label><select id="cl-col">${opts(nums)}</select></div>
    <div class="field"><label>Method</label><select id="cl-method">${['iqr','zscore','manual'].map(m=>`<option>${m}</option>`).join('')}</select></div>
    <div class="field"><label>k</label><input id="cl-k" value="1.5"/></div></div>
    <div class="row"><div class="field"><label>Min (manual)</label><input id="cl-lo"/></div><div class="field"><label>Max (manual)</label><input id="cl-hi"/></div></div>
    <label style="font-size:12px;"><input type="checkbox" id="cl-drop"> drop rows (default: clip to bounds)</label>`;
  else if(op==='rename') h=`<div class="row"><div class="field"><label>Column</label><select id="cl-col">${opts(cols)}</select></div><div class="field"><label>New name</label><input id="cl-to"/></div></div>`;
  else if(op==='dropcols') h=`<div class="field"><label>Columns to drop</label><select id="cl-cols" multiple>${opts(cols)}</select></div>`;
  document.getElementById('cl-fields').innerHTML=h;
}
function applyClean(){
  const op=val('cl-op'); let p={};
  if(op==='fillna') p={cols:msel('cl-cols'),method:val('cl-method'),value:val('cl-val')};
  else if(op==='dropna') p={axis:val('cl-axis'),cols:msel('cl-cols')};
  else if(op==='dropdup') p={cols:msel('cl-cols')};
  else if(op==='cast') p={col:val('cl-col'),dtype:val('cl-dtype'),fmt:val('cl-fmt')||null};
  else if(op==='strclean') p={cols:msel('cl-cols'),ops:[...document.querySelectorAll('.cl-strop:checked')].map(c=>c.value)};
  else if(op==='replace') p={col:val('cl-col'),find:val('cl-find'),repl:val('cl-repl'),regex:chk('cl-regex')};
  else if(op==='outlier') p={col:val('cl-col'),method:val('cl-method'),k:val('cl-k'),lo:val('cl-lo'),hi:val('cl-hi'),drop:chk('cl-drop')};
  else if(op==='rename'){ p={col:val('cl-col'),to:val('cl-to')}; if(!p.to){ toast('Enter a new name',true); return; } }
  else if(op==='dropcols'){ p={cols:msel('cl-cols')}; if(!p.cols.length){ toast('Pick columns to drop',true); return; } }
  runTx(op,p);
}

// ── Feature engineering ──
const FEAT_OPS=[['compute','Computed column (formula)'],['split','Split column'],['explode','Split into rows (explode)'],['bin','Binning'],['encode','Encoding'],['dateparts','Date parts'],['window','Window feature'],['scale','Scale / normalize']];
function openFeature(){ if(!CUR){ toast('Select a table first',true); return; }
  showModal(`<h3>🛠 Feature engineering <span class="muted" style="font-size:11px;">— in-place · undoable</span></h3>
    <div class="field"><label>Operation</label><select id="ft-op" onchange="featureFields()">${FEAT_OPS.map(o=>`<option value="${o[0]}">${o[1]}</option>`).join('')}</select></div>
    <div id="ft-fields"></div>
    <div style="margin-top:12px;"><button class="btn accent" onclick="applyFeature()">▶ Apply</button></div>`);
  featureFields();
}
function insExpr(c){ const e=document.getElementById('ft-expr'); const tok=/^[A-Za-z_]\w*$/.test(c)?c:("col('"+String(c).replace(/'/g,"\\'")+"')"); e.value=(e.value?e.value+' ':'')+tok; e.focus(); }
function splitPreviewNames(){ const c=val('ft-col')||'col'; const n=Math.max(2,parseInt(val('ft-parts'))||2); const into=document.getElementById('ft-into'); if(into) into.placeholder='auto: '+Array.from({length:n},(_,i)=>c+'_'+(i+1)).join(', '); }
function featureFields(){
  const op=val('ft-op'), cols=colsOf(CUR), nums=numColsOf(CUR); let h='';
  if(op==='compute'){
    const chips=cols.map(c=>`<span class="chip" style="cursor:pointer" onclick='insExpr(${JSON.stringify(c)})'>${esc(c)}</span>`).join(' ');
    h=`<div class="field"><label>New column name</label><input id="ft-name" placeholder="revenue"/></div>
       <div class="field"><label>Formula</label><input id="ft-expr" placeholder="qty * price"/></div>
       <div class="muted" style="font-size:11px;">click to insert: ${chips}<br>functions: log ln log10 sqrt abs exp round floor ceil min max pow coalesce · spaces → <code>col('My Col')</code></div>`;
  } else if(op==='split'){
    h=`<div class="row"><div class="field"><label>Column</label><select id="ft-col" onchange="splitPreviewNames()">${opts(cols)}</select></div>
       <div class="field"><label>Split on (delimiter / character)</label><input id="ft-sep" placeholder="e.g.  -  or  ,  or  /" oninput="splitPreviewNames()"/></div>
       <div class="field"><label># of columns</label><input id="ft-parts" value="2" oninput="splitPreviewNames()"/></div></div>
       <div class="field"><label>New column names (comma-separated, optional)</label><input id="ft-into" placeholder="auto: col_1, col_2 …"/></div>
       <label style="font-size:12px;margin-right:14px;"><input type="checkbox" id="ft-remainder" checked> last column keeps the remainder <span class="muted">(e.g. "a-b-c" → "a", "b-c")</span></label>
       <label style="font-size:12px;"><input type="checkbox" id="ft-dropsrc"> drop original column</label>`;
  } else if(op==='explode'){
    h=`<div class="row"><div class="field"><label>Column</label><select id="ft-col">${opts(cols)}</select></div>
       <div class="field"><label>Split on (delimiter / character)</label><input id="ft-sep" placeholder="e.g.  ,  or  ;  (blank if already a list)"/></div></div>
       <label style="font-size:12px;margin-right:14px;"><input type="checkbox" id="ft-strip" checked> trim whitespace around each part</label>
       <label style="font-size:12px;"><input type="checkbox" id="ft-dropempty" checked> drop empty / null rows</label>
       <div class="muted" style="font-size:11px;margin-top:6px;">one row "A, B, C" becomes three rows (A / B / C) — every other column is duplicated.</div>`;
  } else if(op==='bin'){
    h=`<div class="row"><div class="field"><label>Column</label><select id="ft-col">${opts(nums)}</select></div>
       <div class="field"><label>Method</label><select id="ft-method" onchange="document.getElementById('ft-edgewrap').style.display=this.value==='custom'?'':'none';document.getElementById('ft-binwrap').style.display=this.value==='custom'?'none':'';">${['quantile','width','custom'].map(m=>`<option>${m}</option>`).join('')}</select></div>
       <div class="field" id="ft-binwrap"><label># bins</label><input id="ft-bins" value="4"/></div></div>
       <div class="row"><div class="field" id="ft-edgewrap" style="display:none;"><label>Edges (comma)</label><input id="ft-edges" placeholder="0,10,100"/></div>
       <div class="field"><label>New name (optional)</label><input id="ft-newname" placeholder="(col)_bin"/></div></div>`;
  } else if(op==='encode'){
    h=`<div class="row"><div class="field"><label>Column</label><select id="ft-col">${opts(cols)}</select></div>
       <div class="field"><label>Method</label><select id="ft-method">${[['onehot','One-hot (dummies)'],['label','Label / ordinal']].map(m=>`<option value="${m[0]}">${m[1]}</option>`).join('')}</select></div></div>
       <label style="font-size:12px;"><input type="checkbox" id="ft-dropfirst"> drop first category (one-hot)</label>`;
  } else if(op==='dateparts'){
    h=`<div class="field"><label>Date column</label><select id="ft-col">${opts(cols)}</select></div>
       <div class="field"><label>Parts</label><div>${['year','month','day','weekday','quarter','week','hour'].map(p=>`<label style="margin-right:12px;font-size:12px;"><input type="checkbox" class="ft-part" value="${p}" ${(p==='year'||p==='month')?'checked':''}> ${p}</label>`).join('')}</div></div>`;
  } else if(op==='window'){
    h=`<div class="row"><div class="field"><label>Function</label><select id="ft-func">${['cumsum','rolling_mean','rank','lag','lead','pct_change'].map(f=>`<option>${f}</option>`).join('')}</select></div>
       <div class="field"><label>Column</label><select id="ft-col">${opts(nums)}</select></div></div>
       <div class="row"><div class="field"><label>Group by (optional)</label><select id="ft-by" multiple>${opts(cols)}</select></div>
       <div class="field"><label>Order by (optional)</label><select id="ft-order"><option value="">—</option>${opts(cols)}</select></div></div>
       <div class="row"><div class="field"><label>Window (rolling)</label><input id="ft-window" value="3"/></div>
       <div class="field"><label>n (lag/lead)</label><input id="ft-n" value="1"/></div>
       <div class="field"><label>New name (optional)</label><input id="ft-newname"/></div></div>`;
  } else if(op==='scale'){
    h=`<div class="row"><div class="field"><label>Columns</label><select id="ft-cols" multiple>${opts(nums)}</select></div>
       <div class="field"><label>Method</label><select id="ft-method">${[['zscore','Z-score (standardize)'],['minmax','Min-max (0..1)']].map(m=>`<option value="${m[0]}">${m[1]}</option>`).join('')}</select></div></div>
       <label style="font-size:12px;"><input type="checkbox" id="ft-inplace"> replace in place (else add _scaled)</label>`;
  }
  document.getElementById('ft-fields').innerHTML=h;
  if(op==='split') splitPreviewNames();
}
function applyFeature(){
  const op=val('ft-op'); let p={};
  if(op==='compute'){ p={name:val('ft-name'),expr:val('ft-expr')}; if(!p.name||!p.expr){ toast('Name and formula required',true); return; } }
  else if(op==='split'){ const parts=Math.max(2,parseInt(val('ft-parts'))||2); const into=(val('ft-into')||'').split(',').map(s=>s.trim()).filter(Boolean);
    p={col:val('ft-col'),sep:val('ft-sep'),parts,into,remainder:chk('ft-remainder'),drop:chk('ft-dropsrc')};
    if(!p.sep){ toast('Enter a delimiter / character to split on',true); return; } }
  else if(op==='explode') p={col:val('ft-col'),sep:val('ft-sep'),strip:chk('ft-strip'),dropna:chk('ft-dropempty')};
  else if(op==='bin') p={col:val('ft-col'),method:val('ft-method'),bins:val('ft-bins'),edges:(val('ft-edges')||'').split(',').map(s=>s.trim()).filter(Boolean),newname:val('ft-newname')||undefined};
  else if(op==='encode') p={col:val('ft-col'),method:val('ft-method'),drop_first:chk('ft-dropfirst')};
  else if(op==='dateparts'){ p={col:val('ft-col'),parts:[...document.querySelectorAll('.ft-part:checked')].map(c=>c.value)}; if(!p.parts.length){ toast('Pick at least one part',true); return; } }
  else if(op==='window') p={func:val('ft-func'),col:val('ft-col'),by:msel('ft-by'),order:val('ft-order')||undefined,window:val('ft-window'),n:val('ft-n'),newname:val('ft-newname')||undefined};
  else if(op==='scale'){ p={cols:msel('ft-cols'),method:val('ft-method'),inplace:chk('ft-inplace')}; if(!p.cols.length){ toast('Pick columns',true); return; } }
  runTx(op,p);
}

// ── Profile / EDA ──
async function openProfile(){
  if(!CUR){ toast('Select a table first',true); return; }
  let pr;
  try{ pr=await jpost('/api/profile',{table:CUR}); }catch(err){ toast(err.message,true); return; }
  const rowsHtml=pr.columns.map(c=>{
    const stats=c.numeric
      ? `min ${fmtMaybe(c.min)} · max ${fmtMaybe(c.max)} · mean ${fmtMaybe(c.mean)} · med ${fmtMaybe(c.median)} · std ${fmtMaybe(c.std)}`
      : 'top: '+((c.top||[]).map(t=>esc(t[0])+' ('+t[1]+')').join(', ')||'—');
    const flags=[c.constant?'<span class="chip" style="border-color:var(--bad);color:var(--bad);">constant</span>':'', c.high_card?'<span class="chip" style="border-color:var(--accent);color:var(--accent);">high-card</span>':''].join(' ');
    return `<tr><td><b>${esc(c.name)}</b><br><span class="muted" style="font-size:10px;">${esc(c.dtype)}</span></td>
      <td class="num">${c.null_pct}%<br><span class="muted" style="font-size:10px;">${c.nulls}</span></td>
      <td class="num">${c.unique.toLocaleString()}</td>
      <td style="white-space:normal;max-width:420px;">${stats} ${flags}</td></tr>`;
  }).join('');
  const m=document.getElementById('modal'); m.style.width='min(980px,100%)';
  m.innerHTML=`<h3>🔎 Profile — ${esc(CUR)}</h3>
    <div class="muted" style="margin-bottom:8px;">${pr.rows.toLocaleString()} rows · ${pr.cols} columns · ${pr.duplicates.toLocaleString()} duplicate rows
      <button class="minibtn" style="margin-left:10px;" onclick="showCorr()">📊 Correlation heatmap</button></div>
    <div style="overflow:auto;max-height:58vh;border:1px solid var(--border);border-radius:8px;">
      <table><tr><th>Column</th><th class="num">Null %</th><th class="num">Unique</th><th>Summary</th></tr>${rowsHtml}</table></div>
    <div id="corr-out"></div>
    <div style="text-align:right;margin-top:12px;"><button class="btn" onclick="closeModal()">Close</button></div>`;
  document.getElementById('overlay').classList.add('open');
}
async function showCorr(){
  let c;
  try{ c=await jpost('/api/corr',{table:CUR}); }catch(err){ toast(err.message,true); return; }
  const out=document.getElementById('corr-out');
  out.innerHTML='<div id="corr-plot" style="height:'+Math.max(300,c.cols.length*30+140)+'px;margin-top:10px;"></div>';
  const layout={paper_bgcolor:'#161b22',plot_bgcolor:'#161b22',font:{color:'#e6edf3'},margin:{t:10,l:130,r:20,b:130}};
  Plotly.newPlot('corr-plot',[{type:'heatmap',z:c.matrix,x:c.cols,y:c.cols,colorscale:'RdBu',zmid:0,zmin:-1,zmax:1,
    text:c.matrix.map(r=>r.map(v=>v===null?'':(+v).toFixed(2))),texttemplate:'%{text}',hoverongaps:false,
    colorbar:{tickfont:{color:'#e6edf3'},outlinewidth:0}}],layout,{responsive:true,displaylogo:false});
}

// ── Charts: rendered INLINE below the table (no popup) ──
// How many rows the chart engine pulls into the browser. Aggregation happens
// client-side, so this must cover ALL rows of a realistic local table — otherwise
// a branch-ordered table gets head()-truncated and whole categories vanish from
// both the plot and the chart-filter dropdown. Anything beyond this still warns.
const CHART_FETCH_MAX = 1000000;
let _CHART_TABLE = null;     // table whose rows _CHART_DATA currently holds
let _CHART_SIG = '';         // signature of the filter/search/sort the data was fetched under
let _chartCounter = 0;       // unique suffix per inline chart instance
const _editId = {};          // iid -> saved chart id being edited (or null = new)

function _chartSig(){ const s=st(CUR); return JSON.stringify({s:s.search, f:s.filters, o:s.sort}); }

// fetch the current table's rows for the chart engine (re-fetches when the filter changes)
async function ensureChartData(){
  const sig=_chartSig();
  if(_CHART_DATA && _CHART_TABLE===CUR && _CHART_SIG===sig) return true;   // already current
  const s=st(CUR);
  let d;
  try{ d=await jpost('/api/raw',{table:CUR,sort:s.sort,search:s.search,filters:s.filters,limit:CHART_FETCH_MAX}); }
  catch(err){ toast(err.message,true); return false; }
  _CHART_DATA={cols:d.cols, rows:d.rows, numeric_cols:d.numeric_cols}; _CHART_TABLE=CUR; _CHART_SIG=sig;
  if(d.total>d.rows.length) toast('Charting first '+d.rows.length.toLocaleString()+' of '+d.total.toLocaleString()+' filtered rows');
  return true;
}

// the table's filter/search changed → re-fetch filtered rows and re-plot every open chart
async function refreshChartsForFilter(){
  const inst=document.getElementById('chart-instances');
  if(!inst || !inst.children.length) return;       // no charts open → nothing to do
  if(!await ensureChartData()) return;             // re-fetches because the signature changed
  [...inst.children].forEach(w=>{
    const iid=w.id.replace('wrap-','');
    const out=document.getElementById('chartout-'+iid);
    if(out && out.querySelector('[id^="plotly-"]')) generateChart(iid,'tbl');  // only re-plot already-rendered charts
  });
}

// (re)build the chart section for the selected table — renders its saved charts
async function renderChartPanel(){
  const panel=document.getElementById('chartpanel'), inst=document.getElementById('chart-instances');
  if(!panel) return;
  inst.innerHTML=''; _chartCounter=0; _CHART_DATA=null; _CHART_TABLE=null;
  for(const k in _editId) delete _editId[k];
  if(!CUR){ panel.style.display='none'; return; }
  const t=tableMeta(CUR), saved=(t&&t.charts)||[];
  document.getElementById('chartpanel-title').textContent='📊 Charts — '+CUR+(saved.length?(' ('+saved.length+')'):'');
  panel.style.display=saved.length?'':'none';
  if(saved.length){ if(await ensureChartData()) saved.forEach(c=>addChartInstance(c)); }
}

// add a chart builder to the panel. `saved` = {id,title,cfg} to restore, or null for new
function addChartInstance(saved){
  _chartCounter++;
  const iid='c-inst-'+_chartCounter;
  _editId[iid]= saved? saved.id : null;
  chartTypeState[iid]='bar'; labelState[iid]=false;
  chartOptState[iid]={barmode:'group',orient:'v',corners:false,pattern:false,opacity:false,rangeslider:false,annotate:false,condcolor:false,pinned:false};
  chartFilterState[iid]=[]; _cftRowCounter[iid]=0; comboTypeState[iid]={left:'bar',right:'line'}; paletteState[iid]='PayU';
  const wrap=document.createElement('div'); wrap.id='wrap-'+iid; wrap.style.marginBottom='14px';
  wrap.innerHTML = buildChartBuilder('tbl', _CHART_DATA, iid) +
    `<div style="display:flex;gap:8px;align-items:center;margin-top:8px;padding:0 2px;">
       <span class="muted" id="savestatus-${iid}" style="margin-right:auto;font-size:11px;">${saved?('Saved: '+esc(saved.title)):'Unsaved chart — click 💾 to keep it'}</span>
       <button class="btn good" onclick="saveCurrentChart('${iid}')">💾 Save chart</button>
     </div>`;
  document.getElementById('chart-instances').appendChild(wrap);
  if(saved && saved.cfg){
    applyChartCfg(iid, saved.cfg);     // restore controls + render (body stays collapsed; ⚙ Edit to expand)
  }else{
    const body=document.getElementById('builder-body-'+iid); if(body) body.classList.remove('collapsed');
    const fb=document.querySelector('#chartbuilder-'+iid+' .chart-type-btn'); selectChartType(iid,'bar',fb);
  }
  return iid;
}

// toolbar "Add Chart" → fresh builder, expanded, scrolled into view
async function addNewChart(){
  if(!CUR){ toast('Select a table first',true); return; }
  if(!await ensureChartData()) return;
  document.getElementById('chartpanel').style.display='';
  document.getElementById('chartpanel-title').textContent='📊 Charts — '+CUR;
  const iid=addChartInstance(null);
  const body=document.getElementById('builder-body-'+iid); if(body) body.classList.remove('collapsed');
  document.getElementById('wrap-'+iid).scrollIntoView({behavior:'smooth',block:'start'});
}

// sidebar saved-chart link → ensure its table is shown, then scroll to that chart
async function openSavedChart(table, id){
  if(CUR!==table){ selectTable(table); setTimeout(()=>_scrollToSaved(id), 500); return; }
  _scrollToSaved(id);
}
function _scrollToSaved(id){
  const iid=Object.keys(_editId).find(k=>_editId[k]===id);
  if(!iid) return;
  const w=document.getElementById('wrap-'+iid);
  if(w){ w.scrollIntoView({behavior:'smooth',block:'start'}); const b=document.getElementById('builder-body-'+iid); if(b) b.classList.remove('collapsed'); }
}

async function deleteSavedChart(table, id, ev){
  if(ev) ev.stopPropagation();
  try{ await jpost('/api/charts/delete',{table,id}); toast('Chart deleted'); await refreshTables(); if(CUR===table) await renderChartPanel(); }
  catch(err){ toast(err.message,true); }
}

// ── serialize the builder into a saveable config (adapted from db_fw12) ──
function serializeChartCfg(iid){
  const getV=id=>{ const el=document.getElementById(id); return el?el.value:undefined; };
  const ycEl=document.getElementById('ycols-'+iid);
  const ycols=ycEl?Array.from(ycEl.querySelectorAll('input:checked')).map(i=>i.value):[];
  const clEl=document.getElementById('combo-left-cols-'+iid), crEl=document.getElementById('combo-right-cols-'+iid);
  const cleft=clEl?Array.from(clEl.querySelectorAll('input:checked')).map(i=>i.value):[];
  const cright=crEl?Array.from(crEl.querySelectorAll('input:checked')).map(i=>i.value):[];
  const cft=(chartFilterState[iid]||[]).map(r=>({colIdx:r.colIdx,_colName:r._colName,kind:r.kind||'cat',allowed:r.allowed?[...r.allowed]:null,range:r.range||null,enabled:r.enabled!==false}));
  return {
    ctype:chartTypeState[iid]||'bar', opt:Object.assign({},_getOpt(iid)),
    palette:paletteState[iid]||'PayU', combo:Object.assign({},_getComboTypes(iid)), label:!!labelState[iid],
    xcol:getV('xcol-'+iid), colorby:getV('colorbycol-'+iid),
    sort:getV('sort-'+iid), topn:getV('topn-'+iid), xtype:getV('xaxis-type-'+iid),
    yscale:getV('yscale-'+iid), hover:getV('hovermode-'+iid), colormode:getV('colormode-'+iid),
    agg:getV('aggmode-'+iid), bg:getV('chartbg-'+iid), grid:getV('chartgrid-'+iid),
    font:getV('chartfont-'+iid), border:getV('chartborder-'+iid),
    trend:(document.getElementById('trend-'+iid)?Array.from(document.getElementById('trend-'+iid).selectedOptions).map(o=>o.value):[]),
    refline:getV('refline-'+iid),
    title:getV('chart-title-'+iid), desc:getV('chart-desc-'+iid),
    xlabel:getV('xlabel-custom-'+iid), ylabel:getV('ylabel-custom-'+iid), ylabel2:getV('ylabel2-custom-'+iid),
    ycols, cleft, cright, cft
  };
}

// ── apply a saved config back onto the builder, then render (adapted from db_fw12) ──
function applyChartCfg(iid, cfg){
  chartTypeState[iid]=cfg.ctype||'bar';
  chartOptState[iid]=Object.assign({barmode:'group',orient:'v',corners:false,pattern:false,opacity:false,rangeslider:false,annotate:false,condcolor:false,pinned:false}, cfg.opt||{});
  paletteState[iid]=cfg.palette||'PayU';
  comboTypeState[iid]=Object.assign({left:'bar',right:'line'}, cfg.combo||{});
  labelState[iid]=!!cfg.label;
  chartFilterState[iid]=(cfg.cft||[]).map((r,i)=>({rid:iid+'-r'+i,colIdx:r.colIdx,_colName:r._colName,kind:r.kind||'cat',allowed:r.allowed?new Set(r.allowed):null,range:r.range||null,enabled:r.enabled!==false}));
  _cftRowCounter[iid]=(cfg.cft||[]).length;
  const setV=(id,v)=>{ const el=document.getElementById(id); if(el&&v!==undefined&&v!==null) el.value=v; };
  setV('xcol-'+iid,cfg.xcol); setV('colorbycol-'+iid,cfg.colorby);
  setV('sort-'+iid,cfg.sort); setV('topn-'+iid,cfg.topn); setV('xaxis-type-'+iid,cfg.xtype);
  setV('yscale-'+iid,cfg.yscale); setV('hovermode-'+iid,cfg.hover); setV('colormode-'+iid,cfg.colormode);
  setV('aggmode-'+iid,cfg.agg); setV('chartbg-'+iid,cfg.bg); setV('chartgrid-'+iid,cfg.grid);
  setV('chartfont-'+iid,cfg.font); setV('chartborder-'+iid,cfg.border); setV('refline-'+iid,cfg.refline);
  { const _ts=document.getElementById('trend-'+iid); if(_ts){ const _a=Array.isArray(cfg.trend)?cfg.trend:(cfg.trend?[cfg.trend]:[]); const _sel=_a.filter(x=>x&&x!=='none'); Array.from(_ts.options).forEach(o=>o.selected=_sel.includes(o.value)); if(!_sel.length){ const _n=_ts.querySelector('option[value="none"]'); if(_n)_n.selected=true; } } }
  setV('chart-title-'+iid,cfg.title); setV('chart-desc-'+iid,cfg.desc);
  setV('xlabel-custom-'+iid,cfg.xlabel); setV('ylabel-custom-'+iid,cfg.ylabel); setV('ylabel2-custom-'+iid,cfg.ylabel2);
  // selectChartType resets the Y checkboxes/visibility — run it BEFORE restoring selections
  const tbtn=document.querySelector('#chartbuilder-'+iid+' .chart-type-btn[data-ctype="'+(cfg.ctype||'bar')+'"]');
  if(tbtn) selectChartType(iid, cfg.ctype||'bar', tbtn);
  const yc=document.getElementById('ycols-'+iid);
  if(yc&&cfg.ycols&&cfg.ycols.length) yc.querySelectorAll('input').forEach(i=>i.checked=cfg.ycols.includes(i.value));
  const setChecks=(cid,arr)=>{ const c=document.getElementById(cid); if(c&&arr) c.querySelectorAll('input').forEach(i=>i.checked=arr.includes(i.value)); };
  setChecks('combo-left-cols-'+iid,cfg.cleft); setChecks('combo-right-cols-'+iid,cfg.cright);
  ['left','right'].forEach(side=>{
    const grp=document.getElementById('combo-'+side+'-type-'+iid);
    if(grp){ const ct=(comboTypeState[iid][side]||'').toLowerCase(); grp.querySelectorAll('.combo-type-btn').forEach(b=>b.classList.toggle('active', !!ct && b.textContent.toLowerCase().includes(ct))); }
  });
  // sync palette button highlight + labels button
  const prow=document.querySelector('#chartbuilder-'+iid+' .chart-theme-row');
  if(prow){ prow.querySelectorAll('.palette-btn').forEach(b=>b.classList.toggle('active', b.id==='palette-'+(cfg.palette||'PayU')+'-'+iid)); }
  const lbtn=document.getElementById('label-toggle-'+iid);
  if(lbtn){ lbtn.classList.toggle('active', !!cfg.label); lbtn.textContent='🏷 Labels: '+(cfg.label?'ON':'OFF'); }
  _cftRender(iid);
  generateChart(iid, 'tbl');
}

async function saveCurrentChart(iid){
  if(!CUR){ toast('No table selected',true); return; }
  const cfg=serializeChartCfg(iid);
  let title=((document.getElementById('chart-title-'+iid)||{}).value||'').trim();
  if(!title){ const t=tableMeta(CUR); title='Chart '+(((t&&t.charts)?t.charts.length:0)+1); }
  try{
    const r=await jpost('/api/charts/save',{table:CUR,id:_editId[iid]||null,title,cfg});
    _editId[iid]=r.id;
    const ss=document.getElementById('savestatus-'+iid); if(ss) ss.textContent='Saved: '+title;
    toast('Saved: '+title);
    await refreshTables();   // update sidebar links (panel stays as-is)
  }catch(err){ toast(err.message,true); }
}


// ── adapters: feed the ported db_fw12 chart engine from the Polars backend ──
const CHART_MAX_ROWS = 50000;
const COLORBY_MAX_UNIQUE = 30;
let _CHART_DATA = null;                 // {cols, rows, numeric_cols:[idx]} for the open chart
function findTableData(_){ return _CHART_DATA; }
function getState(_){ return {search:'', filters:{}}; }
function applyFilters(_, rows){ return rows; }     // server already filtered
function refreshTable(){}
function buildSidebar(){}
function _snapshotChart(){}
function toggleChartBuilder(iid){ const b=document.getElementById('builder-body-'+iid); if(b) b.classList.toggle('collapsed'); }
function toggleChartOutput(iid){ const o=document.getElementById('chart-output-wrap-'+iid); if(o) o.style.display=(o.style.display==='none'?'':'none'); }
function removeChartInstance(iid){ const w=document.getElementById('wrap-'+iid); if(w) w.remove(); delete _editId[iid]; }

function isPctCol(colName) {
  const name = String(colName).toLowerCase();
  return name.includes('pct') || name.includes('%');
}
function chartFmtNum(val) {
  const n = parseFloat(val);
  if (isNaN(n)) return val; // not a number — return as-is
  if (NUMFMT === "actual") return val;

  const abs = Math.abs(n);

  if (NUMFMT === "k") {
    if (abs < 1000) return val;
    return (n / 1000).toFixed(2) + "K";
  }
  if (NUMFMT === "m") {
    if (abs < 1000) return val;
    if (abs < 1000000) return (n / 1000).toFixed(2) + "K";
    return (n / 1000000).toFixed(2) + "M";
  }
  if (NUMFMT === "b") {
    if (abs < 1000) return val;
    if (abs < 1000000) return (n / 1000).toFixed(2) + "K";
    if (abs < 1000000000) return (n / 1000000).toFixed(2) + "M";
    return (n / 1000000000).toFixed(2) + "B";
  }
  return val;
}
const MONTH_NAMES = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
function _colType(colIdx, rows) {
  const vals = rows.map(r => r[colIdx]).filter(v => v !== null && v !== '' && v !== 'None' && v !== 'null');
  if (vals.length === 0) return 'categorical';
  const sample = vals.slice(0, 30);
  // Month name check
  const monthCount = sample.filter(v => MONTH_NAMES.includes(String(v).toLowerCase().slice(0,3))).length;
  if (monthCount / sample.length > 0.8) return 'categorical';
  // Date check
  const dateCount = sample.filter(v => !isNaN(Date.parse(v))).length;
  if (dateCount / sample.length > 0.8) {
    const unique = new Set(vals.map(v => String(v)));
    return unique.size <= 30 ? 'categorical' : 'date';
  }
  // Numeric check
  const numCount = sample.filter(v => !isNaN(parseFloat(v))).length;
  if (numCount / sample.length > 0.8) return 'numeric';
  return 'categorical';
}
// Returns an ARRAY of overlay lines: [{name, vals, dash?, color?}, ...] (or null).
// Most modes are a single line; bands (stdband/bollinger/minmax) return several.
function _computeTrend(mode, ys) {
  const v = ys.map(y=>isNaN(y)?null:y);
  const idx = v.map((y,i)=>y===null?-1:i).filter(i=>i>=0);
  if (idx.length < 2) return null;
  const n = v.length;
  const BAND = '#60a5fa';
  const flat = val => v.map(()=>val);
  const mean = idx.reduce((a,i)=>a+v[i],0)/idx.length;

  if (mode==='mean')   return [{name:'Mean', vals:flat(mean)}];
  if (mode==='median') {
    const s = idx.map(i=>v[i]).sort((a,b)=>a-b);
    const m = s.length%2 ? s[(s.length-1)/2] : (s[s.length/2-1]+s[s.length/2])/2;
    return [{name:'Median', vals:flat(m)}];
  }
  if (mode==='cummean') {
    let s=0,c=0;
    return [{name:'Cumulative mean', vals:v.map(y=>{ if(y!==null){ s+=y; c++; } return c?s/c:null; })}];
  }
  if (mode==='linear') return [{name:'Linear fit', vals:_polyfit(v, idx, 1)}];
  if (mode==='poly2')  return [{name:'Poly (2)', vals:_polyfit(v, idx, 2)}];
  if (mode==='poly3')  return [{name:'Poly (3)', vals:_polyfit(v, idx, 3)}];
  if (mode==='minmax') {
    const mn = Math.min(...idx.map(i=>v[i])), mx = Math.max(...idx.map(i=>v[i]));
    return [{name:'Max', vals:flat(mx), dash:'dot', color:BAND}, {name:'Min', vals:flat(mn), dash:'dot', color:BAND}];
  }
  if (mode==='stdband') {
    const sd = Math.sqrt(idx.reduce((a,i)=>a+(v[i]-mean)*(v[i]-mean),0)/idx.length);
    return [{name:'Mean', vals:flat(mean)}, {name:'+1σ', vals:flat(mean+sd), dash:'dot', color:BAND}, {name:'-1σ', vals:flat(mean-sd), dash:'dot', color:BAND}];
  }
  if (mode==='bollinger') {
    const w = Math.min(20, Math.max(2, n));
    const mid = _rollmean(v, w), up=[], lo=[];
    for (let i=0;i<n;i++) {
      let s=0,c=0,arr=[];
      for (let j=Math.max(0,i-w+1); j<=i; j++) { if(v[j]!==null){ s+=v[j]; c++; arr.push(v[j]); } }
      const m = c?s/c:null;
      const sd = c?Math.sqrt(arr.reduce((a,x)=>a+(x-m)*(x-m),0)/c):0;
      up.push(m===null?null:m+2*sd); lo.push(m===null?null:m-2*sd);
    }
    return [{name:'BB mid ('+w+')', vals:mid}, {name:'BB upper', vals:up, dash:'dot', color:BAND}, {name:'BB lower', vals:lo, dash:'dot', color:BAND}];
  }
  if (mode.indexOf('ema')===0) {
    const span = parseInt(mode.slice(3))||10, k = 2/(span+1);
    let ema = null;
    const out = v.map(y=>{ if(y===null) return ema; ema = ema===null ? y : (y*k + ema*(1-k)); return ema; });
    return [{name:'EMA('+span+')', vals:out}];
  }
  if (mode.indexOf('ma')===0) {
    const win = parseInt(mode.slice(2))||3;
    return [{name:'MA('+win+')', vals:_rollmean(v, win)}];
  }
  return null;
}
function _rollmean(v, win) {
  return v.map((_,i)=>{ let s=0,c=0; for (let j=Math.max(0,i-win+1); j<=i; j++) { if(v[j]!==null){ s+=v[j]; c++; } } return c?s/c:null; });
}
// Least-squares polynomial fit of given degree over the index positions.
function _polyfit(v, idx, deg) {
  const sp = new Array(2*deg+1).fill(0), sy = new Array(deg+1).fill(0);
  idx.forEach(i=>{ const x=i, y=v[i]; let xp=1; for(let p=0;p<=2*deg;p++){ sp[p]+=xp; xp*=x; } xp=1; for(let p=0;p<=deg;p++){ sy[p]+=y*xp; xp*=x; } });
  const A=[], b=[];
  for (let r=0;r<=deg;r++) { const row=[]; for (let c=0;c<=deg;c++) row.push(sp[r+c]); A.push(row); b.push(sy[r]); }
  const coef = _solve(A, b);
  if (!coef) return null;
  return v.map((_,i)=>{ let x=i, xp=1, val=0; for(let p=0;p<=deg;p++){ val+=coef[p]*xp; xp*=x; } return val; });
}
// Gauss-Jordan solve of a small dense linear system; null if singular.
function _solve(A, b) {
  const n=b.length, M=A.map((r,i)=>r.concat(b[i]));
  for (let col=0; col<n; col++) {
    let piv=col;
    for (let r=col+1; r<n; r++) if (Math.abs(M[r][col])>Math.abs(M[piv][col])) piv=r;
    if (Math.abs(M[piv][col])<1e-12) return null;
    const t=M[col]; M[col]=M[piv]; M[piv]=t;
    for (let r=0; r<n; r++) {
      if (r===col) continue;
      const f=M[r][col]/M[col][col];
      for (let c=col; c<=n; c++) M[r][c]-=f*M[col][c];
    }
  }
  return M.map((r,i)=>r[n]/r[i]);
}
function _trendLabel(mode) { return ({mean:'Mean',median:'Median',linear:'Linear fit',cummean:'Cumulative mean'})[mode]||'Trend'; }
// ── GRAND TOTAL KEYWORDS ──
const TOTAL_KEYWORDS = ["grand total","total","subtotal","sub total","overall","overall total","net total","sum","gt"];

function isGrandTotal(val) {
  return TOTAL_KEYWORDS.includes(String(val).toLowerCase().trim());
}

// ── CHART BUILDER ──
let chartInstanceCount = {};

function buildChartBuilder(tableId, data, instanceId) {
  const cols        = data.cols;
  const numericCols = data.numeric_cols || [];
  const iid         = instanceId;

  return `
  <div class="chart-builder chart-instance" id="chartbuilder-${iid}">
    <div class="chart-builder-header">
      <div class="chart-builder-title">📊 Chart ${iid.split('-inst-')[1] || ''}</div>
      <div style="display:flex;gap:6px;align-items:center;">
        <button class="chart-edit-btn" id="edit-btn-${iid}" onclick="toggleChartBuilder('${iid}')">⚙ Edit</button>
        <button class="chart-edit-btn" id="chart-toggle-btn-${iid}" onclick="toggleChartOutput('${iid}')">▼ Chart</button>
        <button class="chart-instance-remove" onclick="removeChartInstance('${iid}')">✕ Remove</button>
      </div>
    </div>
    <div class="chart-builder-body collapsed" id="builder-body-${iid}">
      <div class="chart-title-row">
        <div>
          <div class="chart-title-label">Chart Title</div>
          <input class="chart-title-input" id="chart-title-${iid}" placeholder="e.g. Bank Wise Success Rate" type="text"/>
        </div>
        <div>
          <div class="chart-title-label">Description</div>
          <input class="chart-title-input" id="chart-desc-${iid}" placeholder="e.g. Top banks by success %" type="text"/>
        </div>
      </div>
      <!-- AXIS LABEL INPUTS -->
      <div class="chart-axis-label-row" id="axis-label-row-${iid}">
        <div class="chart-axis-label-group">
          <div class="chart-axis-label-tag">X Axis Label</div>
          <input class="chart-axis-label-input" id="xlabel-custom-${iid}" placeholder="Auto (from column name)" type="text" oninput="_autoRegenerate('${iid}')"/>
        </div>
        <div class="chart-axis-label-group" id="ylabel-custom-group-${iid}">
          <div class="chart-axis-label-tag">Y Axis Label</div>
          <input class="chart-axis-label-input" id="ylabel-custom-${iid}" placeholder="Auto (from column name)" type="text" oninput="_autoRegenerate('${iid}')"/>
        </div>
        <div class="chart-axis-label-group" id="ylabel2-custom-group-${iid}" style="display:none;">
          <div class="chart-axis-label-tag">Right Y Label</div>
          <input class="chart-axis-label-input" id="ylabel2-custom-${iid}" placeholder="Auto (from column name)" type="text" oninput="_autoRegenerate('${iid}')"/>
        </div>
      </div>

      <div class="chart-type-row">
        <span class="chart-type-label">Chart Type</span>
        <button class="chart-type-btn active" data-ctype="bar"     onclick="selectChartType('${iid}','bar',this)">▊ Bar</button>
        <button class="chart-type-btn"        data-ctype="line"    onclick="selectChartType('${iid}','line',this)">📈 Line</button>
        <button class="chart-type-btn"        data-ctype="area"    onclick="selectChartType('${iid}','area',this)">🏔 Area</button>
        <button class="chart-type-btn"        data-ctype="scatter" onclick="selectChartType('${iid}','scatter',this)">⬤ Scatter</button>
        <button class="chart-type-btn"        data-ctype="pie"     onclick="selectChartType('${iid}','pie',this)">🥧 Pie</button>
        <button class="chart-type-btn"        data-ctype="donut"   onclick="selectChartType('${iid}','donut',this)">🍩 Donut</button>
        <button class="chart-type-btn"        data-ctype="combo"      onclick="selectChartType('${iid}','combo',this)">📊 Combo</button>
        <button class="chart-type-btn"        data-ctype="waterfall"  onclick="selectChartType('${iid}','waterfall',this)">📉 Waterfall</button>
        <button class="chart-type-btn"        data-ctype="funnel"     onclick="selectChartType('${iid}','funnel',this)">🔽 Funnel</button>
        <button class="chart-type-btn"        data-ctype="heatmap"    onclick="selectChartType('${iid}','heatmap',this)">🟩 Heatmap</button>
        <button class="chart-type-btn"        data-ctype="bubble"     onclick="selectChartType('${iid}','bubble',this)">🫧 Bubble</button>
        <button class="chart-type-btn"        data-ctype="box"        onclick="selectChartType('${iid}','box',this)">📦 Box</button>
        <button class="chart-type-btn"        data-ctype="treemap"    onclick="selectChartType('${iid}','treemap',this)">🌳 Treemap</button>
        <button class="chart-type-btn"        data-ctype="sunburst"   onclick="selectChartType('${iid}','sunburst',this)">☀ Sunburst</button>
        <button class="chart-type-btn"        data-ctype="gauge"      onclick="selectChartType('${iid}','gauge',this)">🎯 Gauge</button>
      </div>

      <div class="chart-options-row" id="bar-opts-${iid}">
        <span class="chart-type-label">Bar Mode</span>
        <div class="chart-opt-group">
          <button class="chart-opt-btn active" onclick="setChartOpt('${iid}','barmode','group',this)">Group</button>
          <button class="chart-opt-btn"        onclick="setChartOpt('${iid}','barmode','stack',this)">Stack</button>
          <button class="chart-opt-btn"        onclick="setChartOpt('${iid}','barmode','overlay',this)">Overlay</button>
        </div>
        <div class="chart-divider"></div>
        <span class="chart-type-label">Orient</span>
        <div class="chart-opt-group">
          <button class="chart-opt-btn active" onclick="setChartOpt('${iid}','orient','v',this)">Vertical</button>
          <button class="chart-opt-btn"        onclick="setChartOpt('${iid}','orient','h',this)">Horizontal</button>
        </div>
        <div class="chart-divider"></div>
        <button class="chart-opt-toggle" id="corners-${iid}" onclick="toggleChartOpt('${iid}','corners')">⬜ Corners: OFF</button>
      </div>

      <div class="chart-options-row">
        <span class="chart-type-label">Sort</span>
        <select class="chart-opt-select" id="sort-${iid}">
          <option value="original">Original</option>
          <option value="asc">Ascending</option>
          <option value="desc">Descending</option>
          <option value="az">A → Z</option>
        </select>
        <div class="chart-divider"></div>
        <span class="chart-type-label">Show</span>
        <select class="chart-opt-select" id="topn-${iid}">
          <option value="all">All</option>
          <option value="5">Top 5</option>
          <option value="10">Top 10</option>
          <option value="15">Top 15</option>
          <option value="b5">Bottom 5</option>
        </select>
        <div class="chart-divider"></div>
        <span class="chart-type-label">X Axis</span>
        <select class="chart-opt-select" id="xaxis-type-${iid}" onchange="_autoRegenerate('${iid}')" title="X axis scale — Auto detects category/numeric, override if needed">
          <option value="auto">Auto</option>
          <option value="category">Category</option>
          <option value="numeric">Numeric</option>
        </select>
        <div class="chart-divider"></div>
        <span class="chart-type-label">Y Scale</span>
        <select class="chart-opt-select" id="yscale-${iid}">
          <option value="linear">Linear</option>
          <option value="log">Log</option>
        </select>
        <div class="chart-divider"></div>
        <span class="chart-type-label">Hover</span>
        <select class="chart-opt-select" id="hovermode-${iid}">
          <option value="closest">Closest</option>
          <option value="x unified">Unified</option>
        </select>
      </div>

      <div class="chart-options-row">
        <button class="chart-opt-toggle" id="pattern-${iid}" onclick="toggleChartOpt('${iid}','pattern')">▤ Pattern: OFF</button>
        <div class="chart-divider"></div>
        <span class="chart-type-label">Color</span>
        <select class="chart-opt-select" id="colormode-${iid}">
          <option value="flat">Flat</option>
          <option value="byvalue">By Value</option>
        </select>
        <div class="chart-divider"></div>
        <button class="chart-opt-toggle" id="opacity-${iid}" onclick="toggleChartOpt('${iid}','opacity')">◐ Opacity: OFF</button>
        <div class="chart-divider"></div>
        <button class="chart-opt-toggle" id="rangeslider-${iid}" onclick="toggleChartOpt('${iid}','rangeslider')">↔ Range Slider: OFF</button>
        <div class="chart-divider"></div>
        <button class="chart-opt-toggle" id="annotate-${iid}" onclick="toggleChartOpt('${iid}','annotate')">📌 Annotate: OFF</button>
        <div class="chart-divider"></div>
        <button class="chart-opt-toggle" id="condcolor-${iid}" onclick="toggleChartOpt('${iid}','condcolor')">🔴 Cond. Color: OFF</button>
        <div class="chart-divider"></div>
        <span class="chart-type-label">Ref Line</span>
        <input class="chart-axis-label-input" id="refline-${iid}" placeholder="e.g. 0.90" type="number" step="any" style="width:80px;" oninput="_autoRegenerate('${iid}')"/>
        <div class="chart-divider"></div>
        <span class="chart-type-label" title="Ctrl/Cmd-click to overlay several trends and compare them">Trend ⌄ multi</span>
        <select class="chart-opt-select chart-trend-multi" id="trend-${iid}" multiple size="7" onchange="_autoRegenerate('${iid}')" title="Ctrl/Cmd-click to overlay multiple trends/moving averages/bands and compare. Click 'None' to clear.">
          <optgroup label="—">
            <option value="none" selected>None</option>
          </optgroup>
          <optgroup label="Central tendency">
            <option value="mean">Mean</option>
            <option value="median">Median</option>
            <option value="cummean">Cumulative mean</option>
          </optgroup>
          <optgroup label="Regression fit">
            <option value="linear">Linear fit</option>
            <option value="poly2">Polynomial (deg 2)</option>
            <option value="poly3">Polynomial (deg 3)</option>
          </optgroup>
          <optgroup label="Moving average">
            <option value="ma2">Moving avg (2)</option>
            <option value="ma3">Moving avg (3)</option>
            <option value="ma4">Moving avg (4)</option>
            <option value="ma5">Moving avg (5)</option>
            <option value="ma6">Moving avg (6)</option>
            <option value="ma7">Moving avg (7)</option>
            <option value="ma10">Moving avg (10)</option>
            <option value="ma12">Moving avg (12)</option>
            <option value="ma20">Moving avg (20)</option>
          </optgroup>
          <optgroup label="Exponential moving average">
            <option value="ema3">Exp moving avg (3)</option>
            <option value="ema5">Exp moving avg (5)</option>
            <option value="ema6">Exp moving avg (6)</option>
            <option value="ema10">Exp moving avg (10)</option>
            <option value="ema12">Exp moving avg (12)</option>
            <option value="ema20">Exp moving avg (20)</option>
          </optgroup>
          <optgroup label="Statistical bands">
            <option value="stdband">±1σ band (mean)</option>
            <option value="bollinger">Bollinger bands (20, 2σ)</option>
            <option value="minmax">Min–max envelope</option>
          </optgroup>
        </select>
      </div>
      <div class="chart-axes-row" id="axes-${iid}">
        <div class="chart-axis-group">
          <span class="chart-axis-label" id="xlabel-${iid}">X Axis (Dimension)</span>
          <select class="chart-axis-select" id="xcol-${iid}">
            ${cols.map((c,i) => '<option value="'+i+'">'+String(c)+'</option>').join('')}
          </select>
        </div>
        <div class="chart-axis-group" id="ygroup-${iid}">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;">
            <span class="chart-axis-label" id="ylabel-${iid}">Y Axis (Metrics)</span>
            <div style="display:flex;gap:4px;" id="yaxis-btns-${iid}">
              <button style="padding:2px 8px;border-radius:4px;font-size:10px;font-weight:500;cursor:pointer;border:0.5px solid var(--nav-border);background:var(--tab-bg);color:var(--text2);font-family:'DM Sans',sans-serif;" onclick="selectAllY('${iid}',true)">All</button>
              <button style="padding:2px 8px;border-radius:4px;font-size:10px;font-weight:500;cursor:pointer;border:0.5px solid var(--nav-border);background:var(--tab-bg);color:var(--text2);font-family:'DM Sans',sans-serif;" onclick="selectAllY('${iid}',false)">None</button>
            </div>
          </div>
          <div class="chart-y-checks" id="ycols-${iid}">
            ${numericCols.map(i => '<label class="chart-y-check"><input type="checkbox" value="'+i+'" checked> '+String(cols[i])+'</label>').join('')}
          </div>
        </div>
        <div class="chart-axis-group" id="colorby-group-${iid}">
          <span class="chart-axis-label">Color By <span style="font-size:9px;color:var(--text3);">(optional — groups bars by category)</span></span>
          <select class="chart-axis-select" id="colorbycol-${iid}" onchange="_autoRegenerate('${iid}')">
            <option value="-1">— None —</option>
            ${cols.map((c,i) => '<option value="'+i+'">'+String(c)+'</option>').join('')}
          </select>
        </div>
      </div>

      <!-- FLEXIBLE COMBO PANEL — shown only when Combo selected -->
      <div class="combo-axes-panel" id="combo-panel-${iid}">
        <div class="combo-axis-half">
          <div class="combo-axis-half-label">
            Left Axis
            <span class="combo-axis-badge left">Y1</span>
          </div>
          <div class="combo-type-group" id="combo-left-type-${iid}">
            <button class="combo-type-btn active" onclick="_setComboType('${iid}','left','bar',this)">▊ Bar</button>
            <button class="combo-type-btn"        onclick="_setComboType('${iid}','left','line',this)">📈 Line</button>
            <button class="combo-type-btn"        onclick="_setComboType('${iid}','left','area',this)">🏔 Area</button>
          </div>
          <div class="combo-y-checks" id="combo-left-cols-${iid}" data-iid="${iid}">
            ${numericCols.map(i => '<label class="chart-y-check"><input type="checkbox" value="'+i+'" '+(i===numericCols[0]?'checked':'')+' data-combo-side="left" onchange="_comboColChange(this)"> '+String(cols[i])+'</label>').join('')}
          </div>
        </div>
        <div class="combo-axis-half">
          <div class="combo-axis-half-label">
            Right Axis
            <span class="combo-axis-badge right">Y2</span>
          </div>
          <div class="combo-type-group" id="combo-right-type-${iid}">
            <button class="combo-type-btn"        onclick="_setComboType('${iid}','right','bar',this)">▊ Bar</button>
            <button class="combo-type-btn active" onclick="_setComboType('${iid}','right','line',this)">📈 Line</button>
            <button class="combo-type-btn"        onclick="_setComboType('${iid}','right','area',this)">🏔 Area</button>
          </div>
          <div class="combo-y-checks" id="combo-right-cols-${iid}" data-iid="${iid}">
            ${numericCols.map(i => '<label class="chart-y-check"><input type="checkbox" value="'+i+'" '+(i===numericCols[1]||numericCols.length===1?'checked':'')+' data-combo-side="right" onchange="_comboColChange(this)"> '+String(cols[i])+'</label>').join('')}
          </div>
        </div>
      </div>
      <!-- CHART FILTER — MULTI ROW (type-aware: category checklist / numeric range) -->
      <div id="cft-container-${iid}">
        <div class="cft-hdr"><span class="chart-type-label">🔎 Chart filters</span>
          <span class="muted" style="font-size:10px;">filter the graph on any column — updates live</span>
          <button class="chart-opt-toggle" id="cft-clear-${iid}" style="display:none;margin-left:auto;" onclick="_cftClearAll('${iid}')">Clear all</button>
        </div>
        <div id="cft-rows-${iid}"></div>
        <div style="margin-bottom:10px;">
          <button class="chart-opt-toggle" onclick="_cftAddRow('${iid}')">+ Add Filter</button>
        </div>
      </div>

      <!-- CHART THEME ROW -->
      <div class="chart-theme-row">
        <span class="chart-theme-label">Palette</span>
        <button class=\"palette-btn active\" id=\"palette-PayU-${iid}\" title=\"PayU\" style=\"background:linear-gradient(135deg,#4f8ef7 50%,#22d3a5 50%);\" onclick=\"_setPalette('${iid}','PayU',this)\"></button>
        <button class=\"palette-btn\" id=\"palette-Vibrant-${iid}\" title=\"Vibrant\" style=\"background:linear-gradient(135deg,#ff3366 50%,#ff9500 50%);\" onclick=\"_setPalette('${iid}','Vibrant',this)\"></button>
        <button class=\"palette-btn\" id=\"palette-Pastel-${iid}\" title=\"Pastel\" style=\"background:linear-gradient(135deg,#a8d8ea 50%,#aa96da 50%);\" onclick=\"_setPalette('${iid}','Pastel',this)\"></button>
        <button class=\"palette-btn\" id=\"palette-Mono-${iid}\" title=\"Mono\" style=\"background:linear-gradient(135deg,#e2e8f0 50%,#94a3b8 50%);\" onclick=\"_setPalette('${iid}','Mono',this)\"></button>
        <button class=\"palette-btn\" id=\"palette-Sunset-${iid}\" title=\"Sunset\" style=\"background:linear-gradient(135deg,#ff6b6b 50%,#ffa500 50%);\" onclick=\"_setPalette('${iid}','Sunset',this)\"></button>
        <button class=\"palette-btn\" id=\"palette-Ocean-${iid}\" title=\"Ocean\" style=\"background:linear-gradient(135deg,#0077b6 50%,#00b4d8 50%);\" onclick=\"_setPalette('${iid}','Ocean',this)\"></button>
        <button class=\"palette-btn\" id=\"palette-Forest-${iid}\" title=\"Forest\" style=\"background:linear-gradient(135deg,#2d6a4f 50%,#52b788 50%);\" onclick=\"_setPalette('${iid}','Forest',this)\"></button>
        <div class="chart-divider"></div>
        <span class="chart-theme-label">BG</span>
        <select class="chart-opt-select" id="chartbg-${iid}" onchange="_autoRegenerate('${iid}')">
          <option value="theme">Theme</option>
          <option value="white">White</option>
          <option value="dark">Dark</option>
          <option value="transparent">Transparent</option>
        </select>
        <div class="chart-divider"></div>
        <span class="chart-theme-label">Grid</span>
        <select class="chart-opt-select" id="chartgrid-${iid}" onchange="_autoRegenerate('${iid}')">
          <option value="on">On</option>
          <option value="off">Off</option>
          <option value="h">H Only</option>
        </select>
        <div class="chart-divider"></div>
        <span class="chart-theme-label">Font</span>
        <select class="chart-opt-select" id="chartfont-${iid}" onchange="_autoRegenerate('${iid}')">
          <option value="medium">Medium</option>
          <option value="small">Small</option>
          <option value="large">Large</option>
        </select>
        <div class="chart-divider"></div>
        <span class="chart-theme-label">Border</span>
        <select class="chart-opt-select" id="chartborder-${iid}" onchange="_autoRegenerate('${iid}')">
          <option value="none">None</option>
          <option value="thin">Thin</option>
          <option value="thick">Thick</option>
        </select>
      </div>
      <div class="chart-generate-row">
        <button class="chart-generate-btn" onclick="generateChart('${iid}','${tableId}')">▶ Generate</button>
        <button class="chart-label-toggle" id="label-toggle-${iid}" onclick="toggleLabels('${iid}')">🏷 Labels: OFF</button>
        <button class="chart-opt-toggle" id="pin-${iid}" onclick="_togglePin('${iid}')">📌 Pin: OFF</button>
        <button class="chart-opt-toggle" onclick="_copyChartImage('${iid}')" title="Copy chart as PNG image">📋 Copy Image</button>
        <div class="chart-divider"></div>
        <span class="chart-type-label">Agg</span>
        <select class="chart-opt-select" id="aggmode-${iid}" onchange="_autoRegenerate('${iid}')">
          <option value="none">None</option>
          <option value="sum">Sum</option>
          <option value="avg">Avg</option>
          <option value="count">Count</option>
          <option value="max">Max</option>
          <option value="min">Min</option>
        </select>
        <span style="font-size:11px;color:var(--text3);">Grand Total auto-excluded</span>
      </div>
    </div>
    <div id="chart-output-wrap-${iid}">
      <div id="chart-title-display-${iid}" style="padding:0 16px;"></div>
      <div class="chart-output" id="chartout-${iid}">
        <div class="chart-placeholder">Click ⚙ Edit → configure → Generate</div>
      </div>
    </div>
  </div>`;
}

// ── PALETTE STATE ──
// ── TIME-LIKE COLUMN DETECTION ──
const TIME_KEYWORDS = ['year','month','date','day','week','quarter',
                       'period','fy','fiscal','hour','time','dt','yr',
                       'year_month','yearmonth','yyyymm','mmyyyy'];

function _isTimeCol(colName) {
  const lower = String(colName).toLowerCase();
  return TIME_KEYWORDS.some(kw => lower === kw || lower.includes(kw));
}

// ID-like column names that should be treated as categories even if numeric
const ID_KEYWORDS = ['id','code','mapping','plant','store','sku','no','num','number',
                     'key','ref','pin','zip','postal','branch','region','zone','area',
                     'flag','type','class','cat','seg','grp','group'];

function _isIdCol(colName) {
  const lower = String(colName).toLowerCase().replace(/[_\s-]/g,' ');
  return ID_KEYWORDS.some(kw => lower === kw || lower.split(' ').includes(kw) || lower.endsWith(' ' + kw) || lower.startsWith(kw + ' '));
}

// ── DATE / CHRONOLOGICAL HELPERS (shared: table sort + chart axis) ──
function _isDateStr(v) {
  return /^\d{4}-\d{2}(-\d{2})?$/.test(String(v));
}
function _valsLookDate(vals) {
  const s = (vals||[]).slice(0, 50).map(v => String(v)).filter(v => v && v !== 'null' && v !== 'None');
  if (!s.length) return false;
  return s.filter(_isDateStr).length / s.length > 0.8;
}
function _datesHaveDay(vals) {
  return (vals||[]).slice(0, 50).map(v => String(v)).some(v => /^\d{4}-\d{2}-\d{2}$/.test(v));
}
function _chronoKey(v) {
  const s = String(v);
  if (_isDateStr(s)) return s;
  const M = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
  const mi = M[s.toLowerCase().slice(0,3)];
  if (mi) return 'M' + String(mi).padStart(2,'0');
  const n = parseFloat(s);
  return (isNaN(n) || /^\d{4}-/.test(s)) ? s : n;
}
function _chronoCmp(a, b) {
  const ka = _chronoKey(a), kb = _chronoKey(b);
  return ka < kb ? -1 : ka > kb ? 1 : 0;
}

// ── SMART ROUND FOR AGGREGATED VALUES ──
function _smartRound(v, aggMode) {
  if (v === null || v === undefined || isNaN(v)) return v;
  if (aggMode === 'count') return Math.round(v);
  const abs = Math.abs(v);
  if (abs >= 10000) return Math.round(v);
  if (abs >= 100)   return Math.round(v * 10) / 10;
  if (abs >= 1)     return Math.round(v * 100) / 100;
  return Math.round(v * 10000) / 10000;  // PCT range 0-1
}

const PALETTES = {
  'PayU':    ['#4f8ef7','#22d3a5','#f7b24f','#f75a7a','#a78bfa','#34d399','#60a5fa','#fb923c','#e879f9','#4ade80'],
  'Vibrant': ['#ff3366','#ff9500','#30d5c8','#7b2fff','#00cc44','#ff6600','#e91e63','#2196f3','#4caf50','#ff5722'],
  'Pastel':  ['#a8d8ea','#aa96da','#fcbad3','#ffffd2','#b5ead7','#ffdac1','#c7ceea','#e2f0cb','#ffd3b6','#ffaaa5'],
  'Mono':    ['#e2e8f0','#cbd5e1','#94a3b8','#64748b','#475569','#334155','#1e293b','#0f172a','#f8fafc','#f1f5f9'],
  'Sunset':  ['#ff6b6b','#ffa500','#ffd700','#ff69b4','#ff4500','#dc143c','#ff8c00','#ff1493','#fa8072','#e9967a'],
  'Ocean':   ['#0077b6','#00b4d8','#90e0ef','#48cae4','#023e8a','#0096c7','#caf0f8','#ade8f4','#00b4d8','#0077b6'],
  'Forest':  ['#2d6a4f','#40916c','#52b788','#74c69d','#95d5b2','#b7e4c7','#d8f3dc','#1b4332','#081c15','#40916c'],
};
const paletteState = {};

function _setPalette(iid, name, btn) {
  paletteState[iid] = name;
  const row = btn.closest('.chart-theme-row');
  if (row) row.querySelectorAll('.palette-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  _autoRegenerate(iid);
}

function _getColors(iid) {
  const name = paletteState[iid] || 'PayU';
  return PALETTES[name] || PALETTES['PayU'];
}

// ── CHART FILTER STATE (multi-row) ──
// chartFilterState[iid] = [ {colIdx, allowed:Set|null}, ... ]
const chartFilterState = {};
let   _cftRowCounter   = {};   // per-iid row ID counter

// JSON for inline onclick attrs: escape " as &quot; so it survives a double-quoted attribute
function _qj(s){ return JSON.stringify(s).replace(/"/g,'&quot;'); }

function _getCftRows(iid) {
  if (!chartFilterState[iid]) chartFilterState[iid] = [];
  return chartFilterState[iid];
}

// ── OPEN/CLOSE DROPDOWN TRACKER ──
let _openCftDd = null;
function _closeCftDd() {
  if (_openCftDd) { _openCftDd.remove(); _openCftDd = null; }
}
document.addEventListener('click', e => {
  if (_openCftDd && !_openCftDd.contains(e.target) &&
      !e.target.closest('.chart-filter-trigger') &&
      !e.target.closest('.cft-remove-btn')) _closeCftDd();
});

// ── ADD A FILTER ROW ──  (row: {rid, colIdx, kind:'cat'|'range', allowed:Set|null, range:{min,max}|null, enabled})
function _cftAddRow(iid) {
  const rows = _getCftRows(iid);
  if (rows.length >= 6) { toast('Maximum 6 filters per chart'); return; }
  if (!_cftRowCounter[iid]) _cftRowCounter[iid] = 0;
  const rid = iid + '-r' + (_cftRowCounter[iid]++);
  rows.push({ rid, colIdx: -1, kind: 'cat', allowed: null, range: null, enabled: true });
  _cftRender(iid);
}

function _cftRemoveRow(iid, rid) {
  chartFilterState[iid] = (chartFilterState[iid]||[]).filter(r => r.rid !== rid);
  _cftRender(iid);
  _autoRegenerate(iid);
}

function _cftToggleEnabled(iid, rid) {
  const row = (_getCftRows(iid)).find(r => r.rid === rid);
  if (!row) return;
  row.enabled = row.enabled === false;   // flip (undefined/true -> false, false -> true)
  _cftRender(iid);
  _autoRegenerate(iid);
}

function _cftClearAll(iid) {
  chartFilterState[iid] = [];
  _cftRender(iid);
  _autoRegenerate(iid);
}

// ── RENDER ALL FILTER ROWS ──
function _cftRender(iid) {
  const container = document.getElementById('cft-rows-' + iid);
  if (!container) return;
  const rows = _getCftRows(iid);
  container.innerHTML = rows.map((row, idx) => _cftRowHTML(iid, row, idx)).join('');
  const clr = document.getElementById('cft-clear-' + iid);
  if (clr) clr.style.display = rows.length ? '' : 'none';
}

function _cftRowHTML(iid, row, idx) {
  const colLabel  = row.colIdx >= 0 && row._colName ? row._colName : 'Column...';
  const colActive = row.colIdx >= 0;
  const enabled   = row.enabled !== false;
  let valLabel, valActive;
  if (!colActive) { valLabel = 'Select column first'; valActive = false; }
  else if (row.kind === 'range') {
    const r = row.range || {};
    valActive = r.min !== null && r.min !== undefined || r.max !== null && r.max !== undefined;
    valLabel = valActive
      ? ((r.min !== null && r.min !== undefined ? '≥ ' + r.min : '') + (r.min != null && r.max != null ? '  ·  ' : '') + (r.max !== null && r.max !== undefined ? '≤ ' + r.max : ''))
      : 'Any value';
  } else {
    valActive = row.allowed !== null;
    valLabel = !valActive ? 'All values'
      : [...row.allowed].slice(0,3).join(', ') + (row.allowed.size > 3 ? ' +' + (row.allowed.size-3) + ' more' : '');
  }
  const valClick = row.kind === 'range' ? '_cftOpenRangeDd' : '_cftOpenValDd';

  return '<div class="cft-multi-row' + (enabled?'':' cft-off') + '" id="cft-row-' + row.rid + '">' +
    '<button class="cft-eye" title="Enable / disable this filter" onclick="_cftToggleEnabled(' + _qj(iid) + ',' + _qj(row.rid) + ')">' + (enabled?'👁':'🚫') + '</button>' +
    '<span class="cft-multi-row-num">' + (idx+1) + '</span>' +
    '<div class="cft-multi-group">' +
      '<div class="chart-filter-trigger' + (colActive?' active':'') + '" onclick="_cftOpenColDd(' + _qj(iid) + ',' + _qj(row.rid) + ',this)">' +
        '<span class="chart-filter-trigger-text">' + colLabel + (colActive ? ' <span class="cft-badge">'+(row.kind==='range'?'123':'Aa')+'</span>' : '') + '</span><span>▾</span>' +
      '</div>' +
    '</div>' +
    '<div class="cft-multi-group" style="' + (!colActive ? 'opacity:0.4;pointer-events:none;' : '') + '">' +
      '<div class="chart-filter-trigger' + (valActive?' active':'') + '" onclick="' + valClick + '(' + _qj(iid) + ',' + _qj(row.rid) + ',this)">' +
        '<span class="chart-filter-trigger-text">' + valLabel + '</span><span>▾</span>' +
      '</div>' +
    '</div>' +
    '<button class="cft-remove-btn" onclick="_cftRemoveRow(' + _qj(iid) + ',' + _qj(row.rid) + ')">✕</button>' +
  '</div>';
}

// ── OPEN COLUMN DROPDOWN (all column types; tagged Aa / 123 / 📅) ──
function _cftOpenColDd(iid, rid, triggerEl) {
  _closeCftDd();
  const tableId   = iid.split('-inst-')[0];
  const tableData = findTableData(tableId);
  if (!tableData) return;

  const xColIdx = parseInt((document.getElementById('xcol-' + iid)||{}).value ?? -1);
  const cols    = tableData.cols;
  const rows    = tableData.rows;

  const usedIdxs = new Set(
    (_getCftRows(iid)).filter(r => r.rid !== rid && r.colIdx >= 0).map(r => r.colIdx)
  );

  const cand = cols.map((col, i) => ({col: String(col), idx: i, t: _colType(i, rows)}))
    .filter(o => o.idx !== xColIdx && !usedIdxs.has(o.idx));

  const dd    = document.createElement('div');
  dd.className = 'chart-filter-dd';
  const rect   = triggerEl.getBoundingClientRect();
  dd.style.top  = (rect.bottom + 4) + 'px';
  dd.style.left = rect.left + 'px';
  _openCftDd = dd;

  dd.innerHTML =
    '<input class="chart-filter-dd-search" placeholder="Search columns..." oninput="_cftColSearch(this)">' +
    '<div class="chart-filter-dd-body">' +
    cand.map(o => {
      const badge = o.t==='numeric' ? '123' : o.t==='date' ? '📅' : 'Aa';
      const kind  = o.t==='numeric' ? 'range' : 'cat';
      return '<div class="chart-filter-dd-item" data-iid="' + iid + '" data-rid="' + rid + '" data-idx="' + o.idx + '" data-kind="' + kind + '" data-col="' + o.col.replace(/"/g,'&quot;') + '" onclick="_cftSelectCol(this)">' +
        '<span class="cft-badge">' + badge + '</span><span>' + o.col + '</span></div>';
    }).join('') +
    (cand.length === 0 ? '<div style="padding:8px 10px;font-size:11px;color:var(--text3);">No more columns</div>' : '') +
    '</div>';

  document.body.appendChild(dd);
  const ddRect = dd.getBoundingClientRect();
  if (ddRect.bottom > window.innerHeight) dd.style.top = (rect.top - ddRect.height - 4) + 'px';
  dd.querySelector('.chart-filter-dd-search').focus();
}

function _cftColSearch(input) {
  input.closest('.chart-filter-dd').querySelectorAll('.chart-filter-dd-item').forEach(el => {
    el.style.display = el.textContent.toLowerCase().includes(input.value.toLowerCase()) ? '' : 'none';
  });
}

function _cftSelectCol(el) {
  const iid    = el.dataset.iid;
  const rid    = el.dataset.rid;
  const colIdx = parseInt(el.dataset.idx);
  const colName = el.dataset.col;
  const kind   = el.dataset.kind || 'cat';
  _closeCftDd();

  const row = (_getCftRows(iid)).find(r => r.rid === rid);
  if (!row) return;
  if (row.colIdx !== colIdx) { row.allowed = null; row.range = null; }   // reset on column change
  row.colIdx   = colIdx;
  row._colName = colName;
  row.kind     = kind;
  _cftRender(iid);
  _autoRegenerate(iid);
}

// ── NUMERIC RANGE EDITOR (min/max) ──
function _cftFmtR(v){ return (typeof v==='number' && isFinite(v)) ? (Math.abs(v)>=1000 ? Math.round(v).toLocaleString() : Math.round(v*100)/100) : v; }
function _cftOpenRangeDd(iid, rid, triggerEl) {
  _closeCftDd();
  const row = (_getCftRows(iid)).find(r => r.rid === rid);
  if (!row || row.colIdx < 0) return;
  const tableData = findTableData(iid.split('-inst-')[0]);
  if (!tableData) return;
  const nums = tableData.rows.map(r => parseFloat(r[row.colIdx])).filter(v => !isNaN(v));
  const dMin = nums.length ? Math.min(...nums) : 0, dMax = nums.length ? Math.max(...nums) : 0;
  const cur = row.range || {};

  const dd = document.createElement('div');
  dd.className = 'chart-filter-dd'; dd.dataset.iid = iid; dd.dataset.rid = rid;
  const rect = triggerEl.getBoundingClientRect();
  dd.style.top = (rect.bottom + 4) + 'px'; dd.style.left = rect.left + 'px';
  _openCftDd = dd;

  dd.innerHTML =
    '<div style="padding:8px 10px 4px;font-size:11px;color:var(--text3);">Data range: ' + _cftFmtR(dMin) + ' to ' + _cftFmtR(dMax) + '</div>' +
    '<div class="cft-range-row"><label>Min</label><input type="number" step="any" id="cft-min-' + rid + '" value="' + (cur.min!=null?cur.min:'') + '" placeholder="' + _cftFmtR(dMin) + '"></div>' +
    '<div class="cft-range-row"><label>Max</label><input type="number" step="any" id="cft-max-' + rid + '" value="' + (cur.max!=null?cur.max:'') + '" placeholder="' + _cftFmtR(dMax) + '"></div>' +
    '<div class="chart-filter-dd-footer"><span class="chart-filter-count">blank = open end</span><div style="display:flex;gap:6px;">' +
      '<button class="chart-filter-clear" onclick="_cftRangeClear(this)">Clear</button>' +
      '<button class="chart-filter-apply" onclick="_cftRangeApply(this)">Apply</button></div></div>';

  document.body.appendChild(dd);
  const ddRect = dd.getBoundingClientRect();
  if (ddRect.bottom > window.innerHeight) dd.style.top = (rect.top - ddRect.height - 4) + 'px';
  dd.querySelector('input').focus();
}
function _cftRangeApply(btn) {
  const dd = btn.closest('.chart-filter-dd'), iid = dd.dataset.iid, rid = dd.dataset.rid;
  const mnEl = document.getElementById('cft-min-' + rid), mxEl = document.getElementById('cft-max-' + rid);
  const mn = mnEl.value === '' ? null : parseFloat(mnEl.value);
  const mx = mxEl.value === '' ? null : parseFloat(mxEl.value);
  const row = (_getCftRows(iid)).find(r => r.rid === rid);
  if (row) row.range = (mn === null && mx === null) ? null : { min: mn, max: mx };
  _closeCftDd(); _cftRender(iid); _autoRegenerate(iid);
}
function _cftRangeClear(btn) {
  const dd = btn.closest('.chart-filter-dd'), iid = dd.dataset.iid, rid = dd.dataset.rid;
  const row = (_getCftRows(iid)).find(r => r.rid === rid);
  if (row) row.range = null;
  _closeCftDd(); _cftRender(iid); _autoRegenerate(iid);
}

// ── OPEN VALUES DROPDOWN ──
async function _cftOpenValDd(iid, rid, triggerEl) {
  _closeCftDd();
  const row = (_getCftRows(iid)).find(r => r.rid === rid);
  if (!row || row.colIdx < 0) return;

  const tableId   = iid.split('-inst-')[0];
  const tableData = findTableData(tableId);
  if (!tableData) return;

  // True distinct values come from the server (scans ALL rows), not _CHART_DATA
  // which only holds the first chunk of /api/raw — with branch-ordered data that
  // chunk may contain only a couple of categories. Falls back to the in-memory
  // sample if the request fails.
  const colName = row._colName || String(tableData.cols[row.colIdx]);
  let allVals;
  try {
    const data = await jpost('/api/values', {table: CUR, col: colName, limit: 100000});
    allVals = (data.values || []).map(v => String(v)).filter(v => v && v !== 'null' && v !== 'None').sort();
  } catch (err) {
    allVals = [...new Set(tableData.rows.map(r => String(r[row.colIdx])).filter(v => v && v !== 'null' && v !== 'None'))].sort();
  }
  const allowed = row.allowed || new Set(allVals);
  const allChecked = allowed.size === allVals.length;

  const dd    = document.createElement('div');
  dd.className = 'chart-filter-dd';
  dd.dataset.iid = iid;
  dd.dataset.rid = rid;
  const rect   = triggerEl.getBoundingClientRect();
  dd.style.top  = (rect.bottom + 4) + 'px';
  dd.style.left = rect.left + 'px';
  _openCftDd = dd;

  dd.innerHTML =
    '<input class="chart-filter-dd-search" placeholder="Search values..." oninput="_cftValSearch(this)">' +
    '<div class="chart-filter-dd-body" id="cft-vbody-' + rid + '">' +
    '<div class="chart-filter-dd-item">' +
      '<input type="checkbox" id="cft-all-' + rid + '" ' + (allChecked ? 'checked' : '') + ' onchange="_cftToggleAll(this)">' +
      '<label for="cft-all-' + rid + '" style="font-weight:600;cursor:pointer;">Select All</label>' +
    '</div>' +
    allVals.map(v =>
      '<div class="chart-filter-dd-item" data-val="' + v.replace(/"/g,'&quot;') + '">' +
        '<input type="checkbox" value="' + v.replace(/"/g,'&quot;') + '" ' + (allowed.has(v) ? 'checked' : '') + ' onchange="_cftValCheck(this)">' +
        '<label style="cursor:pointer;">' + v + '</label>' +
      '</div>'
    ).join('') +
    '</div>' +
    '<div class="chart-filter-dd-footer">' +
      '<span class="chart-filter-count" id="cft-vcount-' + rid + '">' + allowed.size + ' of ' + allVals.length + ' selected</span>' +
      '<div style="display:flex;gap:6px;">' +
        '<button class="chart-filter-clear" onclick="_cftValClear(this)">Clear</button>' +
        '<button class="chart-filter-apply" onclick="_cftValApply(this)">Apply</button>' +
      '</div>' +
    '</div>';

  document.body.appendChild(dd);
  const ddRect = dd.getBoundingClientRect();
  if (ddRect.bottom > window.innerHeight) dd.style.top = (rect.top - ddRect.height - 4) + 'px';
  dd.querySelector('.chart-filter-dd-search').focus();
}

function _cftValSearch(input) {
  const body = input.closest('.chart-filter-dd').querySelector('[id^="cft-vbody-"]');
  if (!body) return;
  const q = input.value.toLowerCase();
  body.querySelectorAll('.chart-filter-dd-item[data-val]').forEach(el => {
    el.style.display = el.dataset.val.toLowerCase().includes(q) ? '' : 'none';
  });
}

function _cftToggleAll(cb) {
  const body = cb.closest('.chart-filter-dd').querySelector('[id^="cft-vbody-"]');
  if (!body) return;
  body.querySelectorAll('input[type=checkbox][value]').forEach(c => c.checked = cb.checked);
  _cftUpdateCount(cb.closest('.chart-filter-dd'));
}

function _cftValCheck(cb) {
  const dd    = cb.closest('.chart-filter-dd');
  const body  = dd.querySelector('[id^="cft-vbody-"]');
  const allCb = dd.querySelector('[id^="cft-all-"]');
  if (body && allCb) allCb.checked = [...body.querySelectorAll('input[value]')].every(c => c.checked);
  _cftUpdateCount(dd);
}

function _cftUpdateCount(dd) {
  const body  = dd.querySelector('[id^="cft-vbody-"]');
  const count = dd.querySelector('[id^="cft-vcount-"]');
  if (!body || !count) return;
  const total   = body.querySelectorAll('input[value]').length;
  const checked = body.querySelectorAll('input[value]:checked').length;
  count.textContent = checked + ' of ' + total + ' selected';
}

function _cftValApply(btn) {
  const dd   = btn.closest('.chart-filter-dd');
  const iid  = dd.dataset.iid;
  const rid  = dd.dataset.rid;
  const body = dd.querySelector('[id^="cft-vbody-"]');
  const checked = new Set([...body.querySelectorAll('input[value]:checked')].map(cb => cb.value));
  const row = (_getCftRows(iid)).find(r => r.rid === rid);
  if (row) row.allowed = checked.size > 0 ? checked : null;
  _closeCftDd();
  _cftRender(iid);
  _autoRegenerate(iid);
}

function _cftValClear(btn) {
  const dd  = btn.closest('.chart-filter-dd');
  const iid = dd.dataset.iid;
  const rid = dd.dataset.rid;
  const row = (_getCftRows(iid)).find(r => r.rid === rid);
  if (row) row.allowed = null;
  _closeCftDd();
  _cftRender(iid);
  _autoRegenerate(iid);
}

// ── APPLY ALL CHART FILTERS TO ROWS (enabled rows only; cat = membership, range = min/max) ──
function _applyChartFilter(iid, rows) {
  const filters = _getCftRows(iid).filter(r =>
    r.enabled !== false && r.colIdx >= 0 &&
    ((r.kind === 'range' && r.range) || (r.kind !== 'range' && r.allowed !== null))
  );
  if (filters.length === 0) return rows;
  return rows.filter(row => filters.every(flt => {
    if (flt.kind === 'range') {
      const v = parseFloat(row[flt.colIdx]);
      if (isNaN(v)) return false;
      if (flt.range.min !== null && flt.range.min !== undefined && v < flt.range.min) return false;
      if (flt.range.max !== null && flt.range.max !== undefined && v > flt.range.max) return false;
      return true;
    }
    return flt.allowed.has(String(row[flt.colIdx]));
  }));
}
// ══════════════════════════════════════════════
// BATCH 2 FEATURES
// ══════════════════════════════════════════════

// ── COLUMN RESIZE ──
let _resizeState = null;
function _colResizeStart(e, colIdx, tableId) {
  e.preventDefault(); e.stopPropagation();
  const th = e.target.closest('th');
  if (!th) return;
  _resizeState = { th, startX: e.clientX, startW: th.offsetWidth, colIdx, tableId };
  document.addEventListener('mousemove', _colResizeMove);
  document.addEventListener('mouseup',   _colResizeEnd);
}
function _colResizeMove(e) {
  if (!_resizeState) return;
  const w = Math.max(50, _resizeState.startW + e.clientX - _resizeState.startX);
  _resizeState.th.style.minWidth = w + 'px';
  _resizeState.th.style.width    = w + 'px';
}
function _colResizeEnd() {
  _resizeState = null;
  document.removeEventListener('mousemove', _colResizeMove);
  document.removeEventListener('mouseup',   _colResizeEnd);
}

// ── STATISTICAL SUMMARY PANEL ──
const _statPanelState = {};  // tableId → visible

function toggleStatPanel(tableId) {
  _statPanelState[tableId] = !_statPanelState[tableId];
  const panel = document.getElementById('stat-panel-' + tableId);
  if (!panel) return;
  panel.classList.toggle('visible', !!_statPanelState[tableId]);
  if (_statPanelState[tableId]) _buildStatPanel(tableId);
}

function _buildStatPanel(tableId) {
  const panel = document.getElementById('stat-panel-' + tableId);
  if (!panel) return;
  const data = findTableData(tableId);
  if (!data) return;
  const numCols = (data.numeric_cols || []).slice(0, 8);  // max 8 cols
  if (numCols.length === 0) { panel.innerHTML = '<div style="color:var(--text3);font-size:11px;">No numeric columns</div>'; return; }

  panel.innerHTML = numCols.map(ci => {
    const vals = data.rows.map(r => parseFloat(r[ci])).filter(v => !isNaN(v));
    if (vals.length === 0) return '';
    const sum  = vals.reduce((a,b)=>a+b,0);
    const mean = sum / vals.length;
    const sorted = [...vals].sort((a,b)=>a-b);
    const median = sorted.length % 2 === 0
      ? (sorted[sorted.length/2-1] + sorted[sorted.length/2]) / 2
      : sorted[Math.floor(sorted.length/2)];
    const variance = vals.reduce((a,v)=>a+(v-mean)**2,0)/vals.length;
    const std = Math.sqrt(variance);
    const colName = String(data.cols[ci]);
    return `<div>
      <div class="stat-col-name">${colName}</div>
      <div class="stat-grid">
        <div class="stat-item"><div class="stat-label">Count</div><div class="stat-value">${vals.length}</div></div>
        <div class="stat-item"><div class="stat-label">Sum</div><div class="stat-value">${chartFmtNum(String(Math.round(sum)))}</div></div>
        <div class="stat-item"><div class="stat-label">Mean</div><div class="stat-value">${mean.toFixed(2)}</div></div>
        <div class="stat-item"><div class="stat-label">Median</div><div class="stat-value">${median.toFixed(2)}</div></div>
        <div class="stat-item"><div class="stat-label">Std Dev</div><div class="stat-value">${std.toFixed(2)}</div></div>
        <div class="stat-item"><div class="stat-label">Min</div><div class="stat-value">${chartFmtNum(String(sorted[0]))}</div></div>
        <div class="stat-item"><div class="stat-label">Max</div><div class="stat-value">${chartFmtNum(String(sorted[sorted.length-1]))}</div></div>
        <div class="stat-item"><div class="stat-label">P25</div><div class="stat-value">${sorted[Math.floor(sorted.length*0.25)].toFixed(2)}</div></div>
      </div>
    </div>`;
  }).join('<hr style="border-color:var(--nav-border);margin:10px 0;">');
}

// ── FILTER PRESETS ──
const _filterPresets = {};  // tableId → [{name, filters}, ...]

function saveFilterPreset(tableId) {
  const st = getState(tableId);
  if (!st.filters || Object.keys(st.filters).length === 0) { toast('No active filters to save'); return; }
  const name = prompt('Preset name:');
  if (!name) return;
  if (!_filterPresets[tableId]) _filterPresets[tableId] = [];
  // Deep copy filters (Sets → Arrays for storage)
  const savedFilters = {};
  for (const [k, flt] of Object.entries(st.filters)) {
    savedFilters[k] = flt.type === 'cat' ? {...flt, allowed: [...flt.allowed]} : {...flt};
  }
  _filterPresets[tableId].push({ name, filters: savedFilters });
  _renderPresets(tableId);
  toast('✅ Preset "' + name + '" saved');
}

function _renderPresets(tableId) {
  const wrap = document.getElementById('preset-wrap-' + tableId);
  if (!wrap) return;
  const presets = _filterPresets[tableId] || [];
  wrap.dataset.tid = tableId;
  wrap.innerHTML = presets.map((p,i) =>
    '<span class="preset-tag" data-tid="' + tableId + '" data-idx="' + i + '" onclick="_applyPresetEl(this)" oncontextmenu="event.preventDefault();_deletePresetEl(this)" title="Click to apply | Right-click to delete">' + p.name + '</span>'
  ).join('');
}

function _applyPreset(tableId, idx) {
  const preset = (_filterPresets[tableId] || [])[idx];
  if (!preset) return;
  const st = getState(tableId);
  st.filters = {};
  for (const [k, flt] of Object.entries(preset.filters)) {
    st.filters[k] = flt.type === 'cat' ? {...flt, allowed: new Set(flt.allowed)} : {...flt};
  }
  st.page = 1;
  refreshTable(tableId);
  const iids = _chartInstancesByTable[tableId] || [];
  iids.forEach(iid => _autoRegenerate(iid));
  toast('Applied preset: ' + preset.name);
}

function _applyPresetEl(el)  { _applyPreset(el.dataset.tid, parseInt(el.dataset.idx)); }
function _deletePresetEl(el) { _deletePreset(el.dataset.tid, parseInt(el.dataset.idx)); }

function _deletePreset(tableId, idx) {
  if (!_filterPresets[tableId]) return;
  _filterPresets[tableId].splice(idx, 1);
  _renderPresets(tableId);
}

// ── COMPUTED COLUMNS ──
const _computedCols = {};  // tableId → [{name, expr}, ...]

function addComputedCol(tableId) {
  const data = findTableData(tableId);
  if (!data) return;
  const colNames = data.cols.join(', ');
  const name = prompt('Column name:');
  if (!name) return;
  const expr = prompt('Formula — use the EXACT column names shown below.\nExample: REVENUE / SALES QUANTITY * 100\n\nAvailable columns:\n' + colNames);
  if (!expr) return;
  if (!_computedCols[tableId]) _computedCols[tableId] = [];
  _computedCols[tableId].push({ name: name, expr: expr });
  refreshTable(tableId);
  _persistSoon();
  toast('✅ Computed column "' + name + '" added');
}

function _removeComputedCol(tableId, idx) {
  if (_computedCols[tableId] && _computedCols[tableId][idx]) {
    const nm = _computedCols[tableId][idx].name;
    _computedCols[tableId].splice(idx, 1);
    refreshTable(tableId);
    _persistSoon();
    toast('Removed computed column "' + nm + '"');
  }
}

function _evalComputedRow(expr, row, cols) {
  try {
    // Map each column to a numeric context slot (c0, c1, …)
    const ctx = {};
    cols.forEach((col, i) => { const n = parseFloat(row[i]); ctx['c' + i] = isNaN(n) ? 0 : n; });
    // Substitute the EXACT column names (longest first to avoid partial overlaps).
    // Two-phase via a null-char placeholder so a column literally named "c"/"ctx"/"c0"
    // can't corrupt the generated "(ctx.cN)" reference text. Column headers never contain a NUL char.
    let e = String(expr);
    const pairs = cols.map((col, i) => ({ name: String(col), i }))
                      .sort((a, b) => b.name.length - a.name.length);
    pairs.forEach(p => { if (p.name) e = e.split(p.name).join('\u0000' + p.i + '\u0000'); });
    pairs.forEach(p => { e = e.split('\u0000' + p.i + '\u0000').join('(ctx.c' + p.i + ')'); });
    const result = new Function('ctx', 'return (' + e + ')')(ctx);
    return (typeof result === 'number' && isFinite(result)) ? Math.round(result * 10000) / 10000 : '—';
  } catch(err) { return '—'; }
}

// ── KEYBOARD SHORTCUTS MODAL ──
const SHORTCUTS = [
  ['E', 'Excel number format'],
  ['R', 'Raw number format'],
  ['Esc', 'Exit fullscreen / close modal'],
  ['?', 'Show this shortcuts panel'],
  ['Shift+Click column', 'Multi-column sort'],
  ['Ctrl+Click column', 'Multi-column sort (same as Shift)'],
];

let _shortcutsOpen = false;
function toggleShortcuts() {
  _shortcutsOpen = !_shortcutsOpen;
  let modal = document.getElementById('shortcuts-modal');
  if (_shortcutsOpen) {
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'shortcuts-modal';
      modal.className = 'shortcuts-modal';
      modal.onclick = e => { if (e.target === modal) toggleShortcuts(); };
      modal.innerHTML = `
        <div class="shortcuts-box">
          <div class="shortcuts-title">⌨ Keyboard Shortcuts</div>
          ${SHORTCUTS.map(([k,d]) => `<div class="shortcut-row"><span class="shortcut-key">${k}</span><span class="shortcut-desc">${d}</span></div>`).join('')}
          <div style="margin-top:14px;text-align:right;"><button class="action-btn" onclick="toggleShortcuts()">Close</button></div>
        </div>`;
      document.body.appendChild(modal);
    }
    modal.style.display = 'flex';
  } else if (modal) {
    modal.style.display = 'none';
  }
}

// Add ? key listener
document.addEventListener('keydown', e => {
  if (e.key === '?' && !e.target.matches('input,textarea,select')) toggleShortcuts();
  if (e.key === 'Escape' && _shortcutsOpen) toggleShortcuts();
});

// ── CLICK BAR → FILTER TABLE ──
function _onChartClick(iid, tableId, data) {
  if (!data || !data.points || data.points.length === 0) return;
  const pt = data.points[0];
  const xVal = pt.x !== undefined ? String(pt.x) : (pt.label !== undefined ? String(pt.label) : null);
  if (!xVal) return;

  // Find which column is the X axis
  const xColEl = document.getElementById('xcol-' + iid);
  if (!xColEl) return;
  const xColIdx = parseInt(xColEl.value);

  // Apply as a categorical filter on that column
  const st = getState(tableId);
  if (!st.filters) st.filters = {};
  // Toggle: if already filtered to this value, clear; else set
  const existing = st.filters[xColIdx];
  if (existing && existing.type === 'cat' && existing.allowed.size === 1 && existing.allowed.has(xVal)) {
    delete st.filters[xColIdx];
    toast('Filter cleared: ' + xVal);
  } else {
    st.filters[xColIdx] = { type: 'cat', allowed: new Set([xVal]) };
    toast('Filtered to: ' + xVal + ' (click again to clear)');
  }
  st.page = 1;
  refreshTable(tableId);
}

// ── COPY CHART AS IMAGE ──
function _copyChartImage(iid) {
  const plotEl = document.getElementById('plotly-' + iid);
  if (!plotEl) { toast('Generate a chart first'); return; }
  Plotly.toImage(plotEl, {format:'png', width:1200, height:500}).then(dataUrl => {
    fetch(dataUrl).then(r=>r.blob()).then(blob=>{
      const item = new ClipboardItem({'image/png': blob});
      navigator.clipboard.write([item]).then(()=>toast('✅ Chart copied to clipboard')).catch(()=>{
        // Fallback: open in new tab
        const a = document.createElement('a'); a.href=dataUrl; a.download='chart.png'; a.click();
        toast('Chart saved as PNG');
      });
    });
  });
}

// ── PIN CHART ──
function _togglePin(iid) {
  const opts = _getOpt(iid);
  opts.pinned = !opts.pinned;
  const btn = document.getElementById('pin-' + iid);
  if (btn) {
    btn.classList.toggle('active', opts.pinned);
    btn.textContent = opts.pinned ? '📌 Pin: ON' : '📌 Pin: OFF';
  }
  const builder = document.getElementById('chartbuilder-' + iid);
  if (builder) {
    // When pinned — collapse the builder but keep chart visible
    if (opts.pinned) {
      const editBtn = builder.querySelector('[id^="edit-btn-"]');
      if (editBtn) {
        const panel = document.getElementById('chartpanel-' + iid);
        if (panel && panel.style.display !== 'none') editBtn.click();
      }
      builder.style.border = '1px solid var(--accent)';
      builder.style.boxShadow = '0 0 0 2px rgba(79,110,247,0.2)';
    } else {
      builder.style.border = '';
      builder.style.boxShadow = '';
    }
  }
}

// ── FLEXIBLE COMBO STATE ──
const comboTypeState = {};  // comboTypeState[iid] = { left:'bar', right:'line' }

function _getComboTypes(iid) {
  if (!comboTypeState[iid]) comboTypeState[iid] = { left:'bar', right:'line' };
  return comboTypeState[iid];
}

function _comboColChange(input) {
  // Find the iid from the parent container's data-iid attribute
  const container = input.closest('[data-iid]');
  if (container) _autoRegenerate(container.dataset.iid);
}

function _setComboType(iid, side, ctype, btn) {
  _getComboTypes(iid)[side] = ctype;
  const group = document.getElementById('combo-' + side + '-type-' + iid);
  if (group) group.querySelectorAll('.combo-type-btn').forEach(b => b.classList.toggle('active', b === btn));
  _autoRegenerate(iid);
}

function _getComboYCols(iid, side) {
  const container = document.getElementById('combo-' + side + '-cols-' + iid);
  if (!container) return [];
  return Array.from(container.querySelectorAll('input[data-combo-side="' + side + '"]:checked')).map(i => parseInt(i.value));
}

function selectAllY(iid, checked) {
  const ycols = document.getElementById('ycols-' + iid);
  if (!ycols) return;
  ycols.querySelectorAll('input[type="checkbox"]').forEach(inp => inp.checked = checked);
}

// ── LABELS TOGGLE ──
const labelState = {};
function toggleLabels(iid) {
  labelState[iid] = !labelState[iid];
  const btn = document.getElementById('label-toggle-' + iid);
  if (btn) {
    btn.classList.toggle('active', labelState[iid]);
    btn.textContent = '🏷 Labels: ' + (labelState[iid] ? 'ON' : 'OFF');
  }
  // Auto regenerate chart if already generated
  const chartOut = document.getElementById('chartout-' + iid);
  if (chartOut && chartOut.querySelector('[id^="plotly-"]')) {
    const tableId = iid.split('-inst-')[0];
    generateChart(iid, tableId);
  }
}

const chartTypeState = {};

// ── CHART REGISTRY — for sidebar links ──
const chartRegistry = {};           // chartRegistry[tableId] = [{iid, title}]
const _chartInstancesByTable = {};  // _chartInstancesByTable[tableId] = [iid, ...]

function _registerChart(tableId, iid) {
  if (!_chartInstancesByTable[tableId]) _chartInstancesByTable[tableId] = [];
  if (!_chartInstancesByTable[tableId].includes(iid)) _chartInstancesByTable[tableId].push(iid);
}

function _unregisterChart(tableId, iid) {
  if (_chartInstancesByTable[tableId]) {
    _chartInstancesByTable[tableId] = _chartInstancesByTable[tableId].filter(i => i !== iid);
  }
  if (chartRegistry[tableId]) {
    chartRegistry[tableId] = chartRegistry[tableId].filter(e => e.iid !== iid);
  }
  _refreshSidebar();
}

function _updateChartRegistry(tableId, iid) {
  if (!chartRegistry[tableId]) chartRegistry[tableId] = [];
  const titleEl = document.getElementById('chart-title-' + iid);
  const idx     = (_chartInstancesByTable[tableId] || []).indexOf(iid);
  const label   = (titleEl && titleEl.value.trim()) ? titleEl.value.trim() : 'Chart ' + (idx + 1);
  const existing = chartRegistry[tableId].find(e => e.iid === iid);
  if (existing) { existing.title = label; }
  else          { chartRegistry[tableId].push({ iid, title: label }); }
  _refreshSidebar();
}

function _refreshSidebar() {
  buildSidebar();
}

// ── CHART OPTIONS STATE ──
const chartOptState = {};


function _getOpt(iid) {
  if (!chartOptState[iid]) {
    chartOptState[iid] = {
      barmode: 'group', orient: 'v', corners: false,
      pattern: false, opacity: false, rangeslider: false, annotate: false,
      condcolor: false, pinned: false
    };
  }
  return chartOptState[iid];
}

function setChartOpt(iid, key, val, btn) {
  _getOpt(iid)[key] = val;
  if (btn) {
    btn.parentElement.querySelectorAll('.chart-opt-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  }
  _autoRegenerate(iid);
}

function toggleChartOpt(iid, key) {
  const st  = _getOpt(iid);
  st[key]   = !st[key];
  const btn = document.getElementById(key + '-' + iid);
  if (btn) {
    btn.classList.toggle('active', st[key]);
    const label = key.charAt(0).toUpperCase() + key.slice(1);
    const icons = {corners:'⬜',pattern:'▤',opacity:'◐',rangeslider:'↔',annotate:'📌'};
    btn.textContent = (icons[key]||'') + ' ' + label + ': ' + (st[key] ? 'ON' : 'OFF');
  }
  _autoRegenerate(iid);
}

function _autoRegenerate(iid) {
  const chartOut = document.getElementById('chartout-' + iid);
  if (chartOut && chartOut.querySelector('[id^="plotly-"]')) {
    const tableId = iid.split('-inst-')[0];
    generateChart(iid, tableId);
  }
}

// ── UPDATE AXIS LABEL PLACEHOLDERS ──
function _updateAxisPlaceholders(iid, tableData) {
  if (!tableData) return;
  const cols = tableData.cols;
  const xSel = document.getElementById('xcol-' + iid);
  if (!xSel) return;
  const xIdx    = parseInt(xSel.value);
  const xName   = xIdx >= 0 ? String(cols[xIdx]) : '';
  const ycols   = document.getElementById('ycols-' + iid);
  const yIdxs   = ycols ? Array.from(ycols.querySelectorAll('input:checked')).map(i=>parseInt(i.value)) : [];
  const yName   = yIdxs.map(i => String(cols[i])).join(', ') || 'Y Axis';

  const xEl  = document.getElementById('xlabel-custom-'  + iid);
  const yEl  = document.getElementById('ylabel-custom-'  + iid);
  if (xEl && !xEl.value) xEl.placeholder = xName || 'Auto';
  if (yEl && !yEl.value) yEl.placeholder = yName || 'Auto';
}

function selectChartType(iid, ctype, btn) {
  chartTypeState[iid] = ctype;
  const builder = document.getElementById('chartbuilder-' + iid);
  builder.querySelectorAll('.chart-type-btn').forEach(b => b.classList.toggle('active', b === btn));

  const isPie    = ctype === 'pie' || ctype === 'donut';
  const isCombo  = ctype === 'combo';
  const isBar    = ctype === 'bar';
  const isSpecial = ctype === 'waterfall' || ctype === 'funnel' || ctype === 'heatmap';
  document.getElementById('xlabel-' + iid).textContent = isPie ? 'Labels (Dimension)' : 'X Axis (Dimension)';
  document.getElementById('ylabel-' + iid).textContent = isPie ? 'Values (pick one)' : isCombo ? 'Y Axis (pick 2+: 1st=Bar, 2nd=Line)' : 'Y Axis (Metrics)';

  // Show/hide bar options row
  const barOpts = document.getElementById('bar-opts-' + iid);
  if (barOpts) barOpts.style.display = (isBar || isCombo) ? 'flex' : 'none';
  // Special types: hide color by, bar opts already hidden
  if (isSpecial) {
    const cb = document.getElementById('colorby-group-' + iid);
    if (cb) cb.style.display = 'none';
  }

  // Show/hide Color By group (only for bar, line, area, scatter)
  const colorByGroup = document.getElementById('colorby-group-' + iid);
  if (colorByGroup) colorByGroup.style.display = (isPie || isCombo) ? 'none' : '';

  // Show/hide flexible combo panel + standard Y group
  const comboPanel = document.getElementById('combo-panel-' + iid);
  const ygroup     = document.getElementById('ygroup-' + iid);
  if (comboPanel) comboPanel.classList.toggle('visible', isCombo);
  if (ygroup)     ygroup.style.display = isCombo ? 'none' : '';

  const yaxisBtns = document.getElementById('yaxis-btns-' + iid);
  if (yaxisBtns) yaxisBtns.style.display = isPie ? 'none' : 'flex';
  const ycolsGroup = document.getElementById('ycols-' + iid);
  if (ycolsGroup) {
    ycolsGroup.querySelectorAll('input').forEach(inp => {
      if (isPie) {
        inp.type    = 'radio';
        inp.name    = 'ypie-' + iid;
        inp.checked = false;
      } else {
        inp.type    = 'checkbox';
        inp.checked = true;
      }
    });
    if (isPie) {
      const first = ycolsGroup.querySelector('input');
      if (first) first.checked = true;
    }
  }

  // Show/hide axis label row (hidden for Pie/Donut), Right Y only for Combo
  const axisLabelRow = document.getElementById('axis-label-row-'       + iid);
  const ylabel2Grp   = document.getElementById('ylabel2-custom-group-' + iid);
  const ylabelGrp    = document.getElementById('ylabel-custom-group-'  + iid);
  if (axisLabelRow) axisLabelRow.style.display = isPie ? 'none' : 'flex';
  if (ylabel2Grp)   ylabel2Grp.style.display   = isCombo ? '' : 'none';
  if (ylabelGrp)    ylabelGrp.style.display     = '';
}

// ── GENERATE CHART ──
function generateChart(iid, tableId) {
  if (typeof Plotly === 'undefined') {
    document.getElementById('chartout-' + iid).innerHTML =
      `<div class="chart-error" style="margin:20px;padding:16px;">
        ⚠ Plotly not loaded! Ensure internet connection — Plotly loads from CDN.
      </div>`;
    return;
  }

  const tableDataRaw = findTableData(tableId);
  if (!tableDataRaw) return;

  const ctype      = chartTypeState[iid] || 'bar';
  const isPie      = ctype === 'pie' || ctype === 'donut';
  const isCombo    = ctype === 'combo';
  const isWaterfall = ctype === 'waterfall';
  const isFunnel   = ctype === 'funnel';
  const isHeatmap  = ctype === 'heatmap';
  const isBubble   = ctype === 'bubble';
  const isBox      = ctype === 'box';
  const isTreemap  = ctype === 'treemap';
  const isSunburst = ctype === 'sunburst';
  const isGauge    = ctype === 'gauge';
  const isSpecial  = isWaterfall || isFunnel || isHeatmap || isBubble || isBox || isTreemap || isSunburst || isGauge;
  const showLabels = !!labelState[iid];
  const opts       = _getOpt(iid);
  const COLORS     = _getColors(iid);
  const PATTERNS   = ['/','\\','x','-','|','+','.'];

  // ── READ UI OPTIONS ──
  const xColIdx    = parseInt(document.getElementById('xcol-' + iid).value);
  const ycols      = document.getElementById('ycols-' + iid);
  const yColIdxs   = Array.from(ycols.querySelectorAll('input:checked')).map(i => parseInt(i.value));
  const colorByEl  = document.getElementById('colorbycol-' + iid);
  const colorByIdx = colorByEl ? parseInt(colorByEl.value) : -1;
  const useColorBy = !isPie && !isCombo && colorByIdx >= 0 && colorByIdx !== xColIdx;
  const sortMode   = (document.getElementById('sort-'      + iid)||{}).value || 'original';
  const topN       = (document.getElementById('topn-'      + iid)||{}).value || 'all';
  const yScale     = (document.getElementById('yscale-'    + iid)||{}).value || 'linear';
  const hoverMode  = (document.getElementById('hovermode-' + iid)||{}).value || 'closest';
  const colorMode  = (document.getElementById('colormode-' + iid)||{}).value || 'flat';
  const aggMode    = (document.getElementById('aggmode-'   + iid)||{}).value || 'none';
  const chartBg    = (document.getElementById('chartbg-'   + iid)||{}).value || 'theme';
  const chartGrid  = (document.getElementById('chartgrid-' + iid)||{}).value || 'on';
  const chartFont  = (document.getElementById('chartfont-' + iid)||{}).value || 'medium';
  const chartBorderW = (document.getElementById('chartborder-' + iid)||{}).value || 'none';

  if (yColIdxs.length === 0) { toast('Select at least one Y axis column!'); return; }
  // Combo uses its own left/right panels — no check on yColIdxs needed

  // ── FILTERED ROWS (uses column filters + search) ──
  let rows = applyFilters(tableId, tableDataRaw.rows);
  // also apply search filter
  const st = getState(tableId);
  if (st.search) {
    const q = st.search.toLowerCase();
    rows = rows.filter(r => r.some(c => String(c).toLowerCase().includes(q)));
  }

  // Apply chart-level filter (on top of table column filters)
  rows = _applyChartFilter(iid, rows);

  // Strip null X values (Bug 1 fix)
  rows = rows.filter(r => {
    const v = r[xColIdx];
    return v !== null && v !== '' && String(v) !== 'null' && String(v) !== 'None';
  });

  // Filter Grand Total
  rows = rows.filter(row => !isGrandTotal(row[xColIdx]));

  // ── EMPTY GUARD (Bug 9) ──
  if (rows.length === 0) {
    document.getElementById('chartout-' + iid).innerHTML =
      `<div style="padding:40px;text-align:center;color:var(--text3);">
        📭 No data to chart.<br>
        <span style="font-size:11px;">All rows filtered out — adjust your table filters.</span>
      </div>`;
    return;
  }

  // ── AUTO CHRONOLOGICAL SORT for time/date X (only when Sort = Original) ──
  // Category axes keep data order; our upstream groupby can emit month-then-year.
  {
    const _xcn = String(tableDataRaw.cols[xColIdx]);
    const _xsample = rows.map(r => r[xColIdx]);
    if (sortMode === 'original' && (_isTimeCol(_xcn) || _valsLookDate(_xsample))) {
      rows = [...rows].sort((a,b) => _chronoCmp(a[xColIdx], b[xColIdx]));
    }
  }

  // ── AGGREGATION ──
  function _aggregate(rowsIn, xIdx, yIdx, colorIdx) {
    if (aggMode === 'none') return rowsIn;
    const key = r => colorIdx >= 0 ? String(r[xIdx]) + '||' + String(r[colorIdx]) : String(r[xIdx]);
    const grouped = {};
    rowsIn.forEach(r => {
      const k = key(r);
      if (!grouped[k]) grouped[k] = { xVal: r[xIdx], colorVal: colorIdx>=0?r[colorIdx]:null, vals: [] };
      const n = parseFloat(r[yIdx]);
      if (!isNaN(n)) grouped[k].vals.push(n);
    });
    return Object.values(grouped).map(g => {
      const vals = g.vals;
      let yVal = null;
      if (vals.length > 0) {
        let raw;
        if (aggMode === 'sum')   raw = vals.reduce((a,b)=>a+b,0);
        else if (aggMode === 'avg')   raw = vals.reduce((a,b)=>a+b,0)/vals.length;
        else if (aggMode === 'count') raw = vals.length;
        else if (aggMode === 'max')   raw = Math.max(...vals);
        else if (aggMode === 'min')   raw = Math.min(...vals);
        yVal = _smartRound(raw, aggMode);
      }
      const synth = [...rowsIn[0]];
      synth[xIdx] = g.xVal;
      synth[yIdx] = yVal;
      if (colorIdx >= 0) synth[colorIdx] = g.colorVal;
      return synth;
    });
  }

  // ── SORT ──
  if (sortMode !== 'original' && yColIdxs.length > 0) {
    const syi = yColIdxs[0];
    const _isTimeX = _isTimeCol(String(tableDataRaw.cols[xColIdx])) || _valsLookDate(rows.slice(0,50).map(r => r[xColIdx]));
    if (_isTimeX) {
      // Time/date X axis — always sort chronologically; desc reverses order
      rows = [...rows].sort((a,b) => _chronoCmp(a[xColIdx], b[xColIdx]));
      if (sortMode === 'desc') rows.reverse();
    } else {
      if (sortMode === 'asc')  rows = [...rows].sort((a,b)=>(parseFloat(a[syi])||0)-(parseFloat(b[syi])||0));
      if (sortMode === 'desc') rows = [...rows].sort((a,b)=>(parseFloat(b[syi])||0)-(parseFloat(a[syi])||0));
      if (sortMode === 'az')   rows = [...rows].sort((a,b)=>String(a[xColIdx]).localeCompare(String(b[xColIdx])));
    }
  }

  // ── TOP N ──
  if (topN !== 'all') {
    if (topN === 'b5') {
      rows = [...rows].sort((a,b)=>(parseFloat(a[yColIdxs[0]])||0)-(parseFloat(b[yColIdxs[0]])||0)).slice(0,5);
    } else {
      const n = parseInt(topN);
      if (!isNaN(n)) rows = [...rows].sort((a,b)=>(parseFloat(b[yColIdxs[0]])||0)-(parseFloat(a[yColIdxs[0]])||0)).slice(0,n);
    }
  }

  // ── CHART MAX ROWS CAP (Bug 11) — only when NOT aggregating ──
  // With Aggregation ON, rows collapse to unique X (x Color By), so the raw cap is
  // unnecessary AND harmful (it would truncate before _aggregate sums all rows).
  if (aggMode === 'none' && rows.length > CHART_MAX_ROWS) {
    const _tot = rows.length;
    toast('Showing first ' + CHART_MAX_ROWS.toLocaleString() + ' of ' + _tot.toLocaleString() + ' rows. Turn on Aggregation to use all rows, or raise chart_max_rows.');
    rows = rows.slice(0, CHART_MAX_ROWS);
  }

  // Warn Sum on PCT col
  if (aggMode === 'sum') {
    yColIdxs.forEach(yi => {
      if (isPctCol(String(tableDataRaw.cols[yi]))) toast('⚠ Sum on PCT column may not be meaningful — consider Avg');
    });
  }

  let xVals = rows.map(r => r[xColIdx]);
  const _scType = rows.length > 2000 ? 'scattergl' : 'scatter';  // WebGL for big scatters

  // ── THEME COLORS ──
  const cssProp  = k => getComputedStyle(document.body).getPropertyValue(k).trim();
  const themeBg  = cssProp('--bg');
  const text1    = cssProp('--text1');
  const border   = cssProp('--tbl-border');
  const bgMap    = { theme: themeBg, white:'#ffffff', dark:'#0d1117', transparent:'rgba(0,0,0,0)' };
  const bg       = bgMap[chartBg] || themeBg;
  const fontText = chartBg === 'white' ? '#1a1a1a' : chartBg === 'dark' ? '#e2e8f0' : text1;
  const fontSizeMap = { small:10, medium:12, large:14 };
  const fontSize    = fontSizeMap[chartFont] || 12;
  const borderW     = { none:0, thin:0.5, thick:1.5 }[chartBorderW] || 0;
  const gridColor   = chartGrid === 'off' ? 'transparent' : (chartBg === 'white' ? '#e2e8f0' : border);
  const xGridColor  = (chartGrid === 'h' || chartGrid === 'off') ? 'transparent' : gridColor;

  // ── MARKER BASE ──
  function _markerBase(ti, yVals) {
    const m = { color: COLORS[ti % COLORS.length], opacity: opts.opacity ? 0.7 : 1 };
    if (borderW > 0) m.line = { color: fontText, width: borderW };
    if (opts.pattern) m.pattern = { shape: PATTERNS[ti % PATTERNS.length], solidity:0.5 };
    if (colorMode === 'byvalue' && !useColorBy) {
      m.color     = yVals.map(v => v);
      m.colorscale = 'Blues';
      m.showscale  = true;
      m.colorbar   = { tickfont:{ color:fontText }, outlinewidth:0 };
    }
    return m;
  }

  function _textBase(fmtVals) {
    return {
      text:         showLabels ? fmtVals : [],
      texttemplate: showLabels ? '%{text}' : '',
      textposition: 'outside',
      cliponaxis:   false,
      textfont:     { color:fontText, size:fontSize, family:'DM Sans,sans-serif' },
    };
  }

  // ── BUILD TRACES ──
  let traces = [];

  if (useColorBy) {
    const yIdx = yColIdxs[0];
    // Color By — group by X + ColorBy
    let aggRows = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, colorByIdx) : rows;

    // Get unique ColorBy values
    let uniqueColorVals = [...new Set(aggRows.map(r => String(r[colorByIdx])).filter(v => v && v !== 'null' && v !== 'None'))];
    if (uniqueColorVals.length > COLORBY_MAX_UNIQUE) {
      toast(`⚠ Color By has ${uniqueColorVals.length} unique values — capped at ${COLORBY_MAX_UNIQUE}.`);
      uniqueColorVals = uniqueColorVals.slice(0, COLORBY_MAX_UNIQUE);
    }

    // Get unique X values (preserving order from data)
    const uniqueX = [...new Set(aggRows.map(r => r[xColIdx]).filter(v => v !== null && v !== '' && String(v) !== 'null'))];
    const isHoriz = opts.orient === 'h' && ctype === 'bar';

    traces = uniqueColorVals.map((uval, ti) => {
      const yMap = {};
      aggRows.filter(r => String(r[colorByIdx]) === uval).forEach(r => { yMap[r[xColIdx]] = parseFloat(r[yIdx]); });
      const yVals   = uniqueX.map(x => isNaN(yMap[x]) ? null : yMap[x]);
      const fmtVals = yVals.map(v => v !== null ? chartFmtNum(String(v)) : '');
      const base    = { name: String(uval), marker: _markerBase(ti, yVals), ..._textBase(fmtVals) };

      if (ctype === 'bar') {
        if (isHoriz) {
          const t = { ...base, type:'bar', y:uniqueX, x:yVals, orientation:'h' };
          if (!showLabels) { t.text=[]; t.texttemplate=''; }
          return t;
        }
        return { ...base, type:'bar', x:uniqueX, y:yVals };
      }
      if (ctype === 'line')    return { ...base, type:'scatter', mode:showLabels?'lines+markers+text':'lines+markers', x:uniqueX, y:yVals, textposition:'top center' };
      if (ctype === 'scatter') return { ...base, type:_scType, mode:showLabels?'markers+text':'markers', x:uniqueX, y:yVals, textposition:'top center' };
      if (ctype === 'area')    return { ...base, type:'scatter', mode:showLabels?'lines+markers+text':'lines+markers', fill:'tozeroy', x:uniqueX, y:yVals, textposition:'top center' };
      return { ...base, type:'bar', x:uniqueX, y:yVals };
    });

  } else if (isPie) {
    const yIdx  = yColIdxs[0];
    let aggRows = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, -1) : rows;
    aggRows = aggRows.filter(r => r[xColIdx] !== null && String(r[xColIdx]) !== '');
    const yVals = aggRows.map(r => { const v = parseFloat(r[yIdx]); return isNaN(v) ? null : v; });
    traces = [{
      type:'pie', labels:aggRows.map(r=>r[xColIdx]), values:yVals,
      name:String(tableDataRaw.cols[yIdx]), hole:ctype==='donut'?0.45:0,
      textinfo:showLabels?'label+value+percent':'label+percent',
      hoverinfo:'label+value+percent', textposition:'outside', automargin:true,
      marker:{ colors:COLORS, line:{ color:bg, width:2 } }
    }];

  } else if (isCombo) {
    const isHoriz      = opts.orient === 'h';
    const comboTypes   = _getComboTypes(iid);
    const leftYIdxs    = _getComboYCols(iid, 'left');
    const rightYIdxs   = _getComboYCols(iid, 'right');

    if (leftYIdxs.length === 0 && rightYIdxs.length === 0) {
      toast('Select at least one Y column on left or right axis'); return;
    }

    function _buildComboTrace(yIdx, side, subIdx, ctype) {
      let aggRows   = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, -1) : rows;
      const yVals   = aggRows.map(r => { const v=parseFloat(r[yIdx]); return isNaN(v)?null:v; });
      const fmtVals = yVals.map(v => v!==null?chartFmtNum(String(v)):'');
      const colName = String(tableDataRaw.cols[yIdx]);
      const colorIdx = side === 'left' ? subIdx : leftYIdxs.length + subIdx;
      const xData   = aggRows.map(r=>r[xColIdx]);
      const axisKey = side === 'right' ? 'y2' : 'y';

      let t;
      if (ctype === 'bar') {
        t = { type:'bar', name:colName, yaxis:axisKey, marker:_markerBase(colorIdx,yVals), ..._textBase(fmtVals) };
        if (isHoriz) { t.y=xData; t.x=yVals; t.orientation='h'; delete t.yaxis; t.xaxis = side==='right'?'x2':'x'; if(!showLabels){t.text=[];t.texttemplate='';} }
        else         { t.x=xData; t.y=yVals; }
      } else if (ctype === 'area') {
        t = { type:'scatter', mode:showLabels?'lines+markers+text':'lines+markers', fill:'tozeroy', name:colName, yaxis:axisKey,
          marker:{ color:COLORS[colorIdx%COLORS.length], size:6 }, line:{ color:COLORS[colorIdx%COLORS.length], width:2.5 },
          ..._textBase(fmtVals), textposition:'top center' };
        if (isHoriz) { t.y=xData; t.x=yVals; t.xaxis=side==='right'?'x2':'x'; delete t.yaxis; }
        else         { t.x=xData; t.y=yVals; }
      } else {
        // line (default)
        t = { type:'scatter', mode:showLabels?'lines+markers+text':'lines+markers', name:colName, yaxis:axisKey,
          marker:{ color:COLORS[colorIdx%COLORS.length], size:7 }, line:{ color:COLORS[colorIdx%COLORS.length], width:2.5 },
          ..._textBase(fmtVals), textposition:'top center' };
        if (isHoriz) { t.y=xData; t.x=yVals; t.xaxis=side==='right'?'x2':'x'; delete t.yaxis; }
        else         { t.x=xData; t.y=yVals; }
      }
      return t;
    }

    traces = [
      ...leftYIdxs.map((yIdx, si)  => _buildComboTrace(yIdx, 'left',  si, comboTypes.left)),
      ...rightYIdxs.map((yIdx, si) => _buildComboTrace(yIdx, 'right', si, comboTypes.right)),
    ];

  } else if (isTreemap) {
    const yIdx  = yColIdxs[0];
    let aggRows = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, -1) : rows;
    const labels = aggRows.map(r=>String(r[xColIdx]));
    const vals   = aggRows.map(r=>{const v=parseFloat(r[yIdx]);return isNaN(v)?0:Math.abs(v);});
    traces = [{
      type: 'treemap',
      labels: labels,
      parents: labels.map(()=>''),
      values:  vals,
      texttemplate: '%{label}<br>%{value:.3s}',
      hovertemplate: '<b>%{label}</b><br>'+String(tableDataRaw.cols[yIdx])+': %{value}<br>%{percentRoot:.1%}<extra></extra>',
      marker: { colorscale:'Blues', showscale:false },
      textfont: { family:'DM Sans,sans-serif', size:fontSize },
    }];

  } else if (isSunburst) {
    const yIdx  = yColIdxs[0];
    let aggRows = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, -1) : rows;
    const labels = aggRows.map(r=>String(r[xColIdx]));
    const vals   = aggRows.map(r=>{const v=parseFloat(r[yIdx]);return isNaN(v)?0:Math.abs(v);});
    traces = [{
      type: 'sunburst',
      labels: labels,
      parents: labels.map(()=>''),
      values:  vals,
      hovertemplate: '<b>%{label}</b><br>'+String(tableDataRaw.cols[yIdx])+': %{value}<br>%{percentRoot:.1%}<extra></extra>',
      marker: { colors: COLORS },
      textfont: { family:'DM Sans,sans-serif', size:fontSize, color:fontText },
      leaf: { opacity:0.8 },
    }];

  } else if (isGauge) {
    // Gauge: shows first Y value as a speedometer (best for single-row or aggregated single value)
    const yIdx   = yColIdxs[0];
    let aggRows  = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, -1) : rows;
    const gaugeVal = aggRows.length > 0 ? (parseFloat(aggRows[0][yIdx]) || 0) : 0;
    const allVals  = aggRows.map(r=>parseFloat(r[yIdx])||0).filter(v=>!isNaN(v));
    const gaugeMax = Math.max(...allVals) * 1.2 || 100;
    const gaugeMin = Math.min(0, Math.min(...allVals));
    const colName  = String(tableDataRaw.cols[yIdx]);
    traces = [{
      type: 'indicator',
      mode: 'gauge+number+delta',
      value: gaugeVal,
      title: { text: colName, font: { color: fontText, family: 'DM Sans,sans-serif' } },
      number: { font: { color: fontText, family: 'DM Sans,sans-serif' } },
      delta: { reference: gaugeMax * 0.8 },
      gauge: {
        axis: { range:[gaugeMin, gaugeMax], tickfont:{ color:fontText }, tickcolor:gridColor },
        bar:  { color: COLORS[0] },
        bgcolor: bg,
        bordercolor: gridColor,
        steps: [
          { range:[gaugeMin, gaugeMax*0.5],  color:'rgba(247,90,122,0.15)' },
          { range:[gaugeMax*0.5, gaugeMax*0.8], color:'rgba(247,178,79,0.15)' },
          { range:[gaugeMax*0.8, gaugeMax],  color:'rgba(34,211,165,0.15)' },
        ],
        threshold: { line:{ color:COLORS[0], width:4 }, thickness:0.75, value:gaugeMax*0.8 },
      },
    }];

  } else if (isBubble) {
    // Bubble: X=col, Y=col, Size=col (3rd Y column)
    const xIdx  = xColIdx;
    const yIdx  = yColIdxs[0];
    const szIdx = yColIdxs[1] >= 0 ? yColIdxs[1] : yColIdxs[0];
    let aggRows = aggMode !== 'none' ? _aggregate(rows, xIdx, yIdx, -1) : rows;
    const xVals  = aggRows.map(r=>r[xIdx]);
    const yVals  = aggRows.map(r=>{const v=parseFloat(r[yIdx]);return isNaN(v)?null:v;});
    const szVals = aggRows.map(r=>{const v=parseFloat(r[szIdx]);return isNaN(v)?1:Math.abs(v);});
    const szMax  = Math.max(...szVals.filter(v=>v!==null)) || 1;
    traces = [{
      type:'scatter', mode:'markers',
      x: xVals, y: yVals,
      name: String(tableDataRaw.cols[yIdx]),
      marker:{
        size: szVals.map(v=>5+((v/szMax)*45)),
        color: COLORS[0], opacity:0.7,
        line:{ color:bg, width:1 },
        sizemode:'diameter',
      },
      text: showLabels ? xVals.map((x,i)=>String(x)+'<br>'+chartFmtNum(String(yVals[i]))) : undefined,
      hovertemplate:'<b>%{x}</b><br>'+String(tableDataRaw.cols[yIdx])+': %{y}<br>Size: %{marker.size:.1f}<extra></extra>',
    }];

  } else if (isBox) {
    // Box plot: X=category col, Y=numeric — one box per X category
    const yIdx = yColIdxs[0];
    const uniqueX = [...new Set(rows.map(r=>String(r[xColIdx])))];
    traces = uniqueX.map((xv,ti)=>{
      const yVals = rows.filter(r=>String(r[xColIdx])===xv).map(r=>{const v=parseFloat(r[yIdx]);return isNaN(v)?null:v;}).filter(v=>v!==null);
      return {
        type:'box', name:xv, y:yVals,
        marker:{ color:COLORS[ti%COLORS.length] },
        boxmean:true,
      };
    });

  } else if (isWaterfall) {
    const yIdx  = yColIdxs[0];
    let aggRows = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, -1) : rows;
    const yVals = aggRows.map(r => { const v=parseFloat(r[yIdx]); return isNaN(v)?0:v; });
    const xData = aggRows.map(r=>r[xColIdx]);
    traces = [{
      type: 'waterfall',
      x: xData, y: yVals,
      name: String(tableDataRaw.cols[yIdx]),
      orientation: 'v',
      connector: { line: { color: border } },
      increasing:  { marker: { color: '#22d3a5' } },
      decreasing:  { marker: { color: '#f75a7a' } },
      totals:      { marker: { color: COLORS[0] } },
      textposition: showLabels ? 'outside' : 'none',
      text: showLabels ? yVals.map(v => chartFmtNum(String(v))) : [],
      textfont: { color: fontText, size: fontSize, family: 'DM Sans,sans-serif' },
    }];

  } else if (isFunnel) {
    const yIdx  = yColIdxs[0];
    let aggRows = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, -1) : rows;
    const yVals = aggRows.map(r => { const v=parseFloat(r[yIdx]); return isNaN(v)?0:v; });
    const xData = aggRows.map(r=>r[xColIdx]);
    traces = [{
      type: 'funnel',
      y: xData, x: yVals,
      name: String(tableDataRaw.cols[yIdx]),
      textinfo: showLabels ? 'value+percent initial' : 'percent initial',
      textposition: 'inside',
      textfont: { color: '#fff', size: fontSize, family: 'DM Sans,sans-serif' },
      marker: { color: COLORS, line: { width: 1, color: bg } },
      connector: { line: { color: border, dash: 'dot', width: 1 } },
    }];

  } else if (isHeatmap) {
    // Heatmap: X axis = X column, Y axis = second categorical col, Color = first numeric col
    const yIdx   = yColIdxs[0];
    // Find a second categorical column for Y axis (not X axis col)
    const catCols = tableDataRaw.cols.map((col,i)=>i).filter(i => i !== xColIdx && _colType(i, tableDataRaw.rows) === 'categorical');
    const yAxisColIdx = catCols[0] >= 0 ? catCols[0] : -1;
    if (yAxisColIdx < 0) {
      toast('Heatmap needs a second categorical column for the Y axis');
      return;
    }
    let aggRows = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, yAxisColIdx) : rows;
    const uniqueX = [...new Set(aggRows.map(r=>String(r[xColIdx])))];
    const uniqueY = [...new Set(aggRows.map(r=>String(r[yAxisColIdx])))];
    const zMap = {};
    aggRows.forEach(r => {
      const key = String(r[xColIdx]) + '|||' + String(r[yAxisColIdx]);
      const v   = parseFloat(r[yIdx]);
      zMap[key] = isNaN(v) ? null : v;
    });
    const zVals = uniqueY.map(yv => uniqueX.map(xv => zMap[xv+'|||'+yv] ?? null));
    traces = [{
      type: 'heatmap',
      x: uniqueX, y: uniqueY, z: zVals,
      colorscale: 'Blues',
      showscale: true,
      text: showLabels ? zVals.map(row => row.map(v => v !== null ? chartFmtNum(String(v)) : '')) : undefined,
      texttemplate: showLabels ? '%{text}' : undefined,
      hovertemplate: '<b>%{x}</b> × <b>%{y}</b><br>' + String(tableDataRaw.cols[yIdx]) + ': %{z}<extra></extra>',
      colorbar: { tickfont: { color: fontText }, outlinewidth: 0 },
    }];

  } else {
    const isHoriz = opts.orient === 'h' && ctype === 'bar';
    traces = yColIdxs.map((yIdx, ti) => {
      let aggRows = aggMode !== 'none' ? _aggregate(rows, xColIdx, yIdx, -1) : rows;
      const yVals   = aggRows.map(r => { const v=parseFloat(r[yIdx]); return isNaN(v)?null:v; });
      const fmtVals = yVals.map(v => v!==null?chartFmtNum(String(v)):'');
      const base    = { name:String(tableDataRaw.cols[yIdx]), marker:_markerBase(ti,yVals), ..._textBase(fmtVals) };
      const xData   = aggRows.map(r=>r[xColIdx]);
      if (ctype === 'bar') {
        if (isHoriz) {
          const t = { ...base, type:'bar', y:xData, x:yVals, orientation:'h' };
          if (!showLabels) { t.text=[]; t.texttemplate=''; }
          return t;
        }
        return { ...base, type:'bar', x:xData, y:yVals };
      }
      if (ctype === 'line')    return { ...base, type:'scatter', mode:showLabels?'lines+markers+text':'lines+markers', x:xData, y:yVals, textposition:'top center' };
      if (ctype === 'scatter') return { ...base, type:_scType, mode:showLabels?'markers+text':'markers', x:xData, y:yVals, textposition:'top center' };
      if (ctype === 'area')    return { ...base, type:'scatter', mode:showLabels?'lines+markers+text':'lines+markers', fill:'tozeroy', x:xData, y:yVals, textposition:'top center' };
      return { ...base, type:'bar', x:xData, y:yVals };
    });
  }

  // ── LAYOUT ──
  const isHorizLayout = opts.orient === 'h' && (ctype === 'bar' || isCombo);
  const xColName = String(tableDataRaw.cols[xColIdx]);
  const yColName = yColIdxs.map(i => String(tableDataRaw.cols[i])).join(', ');

  // X axis type — read manual override first, then auto-detect
  const xAxisTypeEl  = document.getElementById('xaxis-type-' + iid);
  const xAxisTypeOverride = xAxisTypeEl ? xAxisTypeEl.value : 'auto';

  let xAxisType;
  if (xAxisTypeOverride === 'category') {
    xAxisType = 'category';
  } else if (xAxisTypeOverride === 'numeric') {
    xAxisType = '-';
  } else {
    // Auto mode:
    //  - real date strings (YYYY-MM / YYYY-MM-DD) -> 'date' (Plotly auto-orders + time scale)
    //  - other time-like / ID-like names, or bar/line/area/scatter -> 'category'
    //  - otherwise numeric
    const isBarLineArea = (ctype === 'bar' || ctype === 'line' || ctype === 'area' || ctype === 'scatter');
    const _xLooksDate   = _valsLookDate(xVals);
    if (!isPie && !isHeatmap && !isGauge) {
      if (_xLooksDate && !isHorizLayout) {
        xAxisType = 'date';
      } else if (_isTimeCol(xColName) || _isIdCol(xColName) || isBarLineArea) {
        xAxisType = 'category';
      } else {
        xAxisType = '-';
      }
    } else {
      xAxisType = '-';
    }
  }

  // ── CUSTOM AXIS LABELS (user overrides, fall back to auto col names) ──
  const _xLblEl   = document.getElementById('xlabel-custom-'  + iid);
  const _yLblEl   = document.getElementById('ylabel-custom-'  + iid);
  const _y2LblEl  = document.getElementById('ylabel2-custom-' + iid);
  const xAxisLabel  = (_xLblEl  && _xLblEl.value.trim())  ? _xLblEl.value.trim()  : xColName;
  const yAxisLabel  = (_yLblEl  && _yLblEl.value.trim())  ? _yLblEl.value.trim()  : yColName;
  const y2AxisLabel = (_y2LblEl && _y2LblEl.value.trim()) ? _y2LblEl.value.trim() : null;

  const baseLayout = {
    paper_bgcolor: bg, plot_bgcolor: bg,
    font:   { color:fontText, family:'DM Sans,sans-serif', size:fontSize },
    margin: { t:60, b:isHorizLayout?60:80, l:isHorizLayout?130:70, r:isCombo?70:40 },
    legend: { bgcolor:'transparent', font:{ color:fontText, size:fontSize } },
    hovermode: hoverMode,
  };

  let layout;
  if (isPie) {
    layout = baseLayout;
  } else if (isCombo) {
    const leftCols  = _getComboYCols(iid, 'left');
    const rightCols = _getComboYCols(iid, 'right');
    const leftAutoTitle  = leftCols.map(i  => String(tableDataRaw.cols[i])).join(', ')  || 'Left Axis';
    const rightAutoTitle = rightCols.map(i => String(tableDataRaw.cols[i])).join(', ') || 'Right Axis';
    const leftTitle  = (_yLblEl  && _yLblEl.value.trim())  ? _yLblEl.value.trim()  : leftAutoTitle;
    const rightTitle = (_y2LblEl && _y2LblEl.value.trim()) ? _y2LblEl.value.trim() : rightAutoTitle;
    layout = Object.assign(baseLayout, {
      xaxis:  { title:{ text:isHorizLayout?leftTitle:xAxisLabel }, tickfont:{ color:fontText,size:fontSize }, gridcolor:xGridColor, linecolor:gridColor,
               tickangle: isHorizLayout?0:(xVals&&xVals.some&&xVals.some(v=>String(v).length>8)?-45:-35),
               ticklen:4, type:xAxisType,
               nticks: (xVals && xVals.length > 20) ? 20 : undefined,
               automargin: true },
      yaxis:  { title:{ text:isHorizLayout?xAxisLabel:leftTitle }, tickfont:{ color:fontText,size:fontSize }, gridcolor:gridColor, linecolor:gridColor, type:yScale },
      barmode: opts.barmode,
    });
    if (isHorizLayout) {
      layout.xaxis2 = { title:{ text:rightTitle }, overlaying:'x', side:'top', tickfont:{ color:fontText,size:fontSize }, gridcolor:'transparent' };
    } else {
      layout.yaxis2 = { title:{ text:rightTitle }, overlaying:'y', side:'right', tickfont:{ color:fontText,size:fontSize }, gridcolor:'transparent' };
    }
  } else {
    layout = Object.assign(baseLayout, {
      xaxis: { title:{ text:isHorizLayout?yAxisLabel:xAxisLabel }, tickfont:{ color:fontText,size:fontSize }, gridcolor:xGridColor, linecolor:gridColor,
               // Bug 3: auto-rotate for long labels, increase bottom margin
               tickangle: isHorizLayout?0:(xVals&&xVals.some&&xVals.some(v=>String(v).length>8)?-45:-35),
               ticklen:4, type:xAxisType,
               // Auto-limit visible labels when >20 data points to prevent congestion
               nticks: (xVals && xVals.length > 20) ? 20 : undefined,
               automargin: true },
      yaxis: { title:{ text:isHorizLayout?xAxisLabel:yAxisLabel }, tickfont:{ color:fontText,size:fontSize }, gridcolor:gridColor, linecolor:gridColor, type:(yScale==='log'&&traces.some(t=>(t.y||t.x||[]).some(v=>parseFloat(v)<=0)))?'linear':yScale,  // Bug 12: negative values incompatible with log
               // Bug 2: if all Y values identical, add ±10% padding so chart isn't a flat line
               range: (()=>{ if(isPie||isCombo) return undefined; const ys=traces.flatMap(t=>t.y||t.x||[]).map(v=>parseFloat(v)).filter(v=>!isNaN(v)); if(!ys.length) return undefined; const mn=Math.min(...ys),mx=Math.max(...ys); if(mn===mx){ const pad=Math.abs(mx)*0.1||1; return [mn-pad,mx+pad]; } return undefined; })() },
      barmode: ctype==='bar' ? opts.barmode : 'relative',
    });
  }

  // ── REFERENCE LINE ──
  const refLineEl  = document.getElementById('refline-' + iid);
  const refLineVal = refLineEl ? parseFloat(refLineEl.value) : NaN;
  if (!isNaN(refLineVal) && !isPie && !isHeatmap && !isFunnel) {
    if (!layout.shapes) layout.shapes = [];
    layout.shapes.push({
      type: 'line', xref: 'paper', x0: 0, x1: 1,
      yref: 'y', y0: refLineVal, y1: refLineVal,
      line: { color: '#f7b24f', width: 2, dash: 'dash' },
    });
    if (!layout.annotations) layout.annotations = [];
    layout.annotations.push({
      xref: 'paper', x: 1.01, xanchor: 'left',
      yref: 'y', y: refLineVal,
      text: chartFmtNum(String(refLineVal)),
      showarrow: false,
      font: { color: '#f7b24f', size: 11, family: 'DM Sans,sans-serif' },
    });
  }

  // ── CONDITIONAL COLORS (red/green threshold) ──
  if (opts.condcolor && !isPie && !isHeatmap && !isFunnel && !isNaN(refLineVal)) {
    traces.forEach(trace => {
      if (trace.type === 'bar') {
        const vals = (isHorizLayout ? trace.x : trace.y) || [];
        trace.marker = { ...trace.marker,
          color: vals.map(v => (parseFloat(v)||0) >= refLineVal ? '#22d3a5' : '#f75a7a'),
        };
      }
    });
  }

  // Range slider
  if (opts.rangeslider && !isPie) {
    if (isHorizLayout) layout.yaxis.rangeslider = { visible:true };
    else               { layout.xaxis.rangeslider = { visible:true }; layout.margin.b=40; }
  }

  // Auto annotate (Bug 9 guard: check non-null)
  if (opts.annotate && !isPie && yColIdxs.length > 0) {
    const firstYIdx = yColIdxs[0];
    const yVals     = rows.map(r => parseFloat(r[firstYIdx])).filter(v => !isNaN(v));
    if (yVals.length > 0) {
      const allY  = rows.map(r => { const v=parseFloat(r[firstYIdx]); return isNaN(v)?null:v; });
      const maxIdx = allY.indexOf(Math.max(...yVals));
      const minIdx = allY.indexOf(Math.min(...yVals));
      layout.annotations = [];
      if (maxIdx >= 0) layout.annotations.push({ x:isHorizLayout?allY[maxIdx]:xVals[maxIdx], y:isHorizLayout?xVals[maxIdx]:allY[maxIdx], text:'▲ Max: '+chartFmtNum(String(allY[maxIdx])), showarrow:true, arrowhead:2, arrowcolor:'#22d3a5', font:{ color:'#22d3a5',size:11,family:'DM Sans,sans-serif' }, bgcolor:bg, bordercolor:'#22d3a5', borderwidth:1, borderpad:4 });
      if (minIdx >= 0 && minIdx !== maxIdx) layout.annotations.push({ x:isHorizLayout?allY[minIdx]:xVals[minIdx], y:isHorizLayout?xVals[minIdx]:allY[minIdx], text:'▼ Min: '+chartFmtNum(String(allY[minIdx])), showarrow:true, arrowhead:2, arrowcolor:'#f75a7a', font:{ color:'#f75a7a',size:11,family:'DM Sans,sans-serif' }, bgcolor:bg, bordercolor:'#f75a7a', borderwidth:1, borderpad:4 });
    }
  }

  const config = {
    responsive:true, displaylogo:false, scrollZoom:true,
    toImageButtonOptions:{ format:'png', filename:'chart' },
    modeBarButtonsToRemove:['select2d','lasso2d','autoScale2d']
  };

  // Title + description
  const titleVal = (document.getElementById('chart-title-' + iid)||{}).value?.trim() || '';
  const descVal  = (document.getElementById('chart-desc-'  + iid)||{}).value?.trim() || '';
  const titleDiv = document.getElementById('chart-title-display-' + iid);
  if (titleDiv) {
    titleDiv.innerHTML = '';
    if (titleVal) { const t=document.createElement('div'); t.className='chart-rendered-title'; t.textContent=titleVal; titleDiv.appendChild(t); }
    if (descVal)  { const d=document.createElement('div'); d.className='chart-rendered-desc';  d.textContent=descVal;  titleDiv.appendChild(d); }
  }

  // Update sidebar chart registry
  _updateChartRegistry(tableId, iid);

  // Show filter info badge if filters active
  const outDiv = document.getElementById('chartout-' + iid);
  const filteredCount = rows.length;
  const totalCount    = tableDataRaw.rows.filter(r => !isGrandTotal(r[xColIdx])).length;
  const cftActive = _getCftRows(iid).some(r => r.colIdx >= 0 && r.allowed !== null);
  const cftBadgeParts = _getCftRows(iid)
    .filter(r => r.colIdx >= 0 && r.allowed !== null)
    .map(r => (r._colName||String(tableDataRaw.cols[r.colIdx])) + ' = ' + [...r.allowed].slice(0,3).join(', ') + (r.allowed.size > 3 ? ' +' + (r.allowed.size-3) + ' more' : ''));
  let filterBadge = '';
  if (filteredCount < totalCount || cftActive) {
    filterBadge = '<div style="font-size:10px;color:var(--accent);padding:2px 12px 4px;display:flex;gap:8px;flex-wrap:wrap;">';
    if (filteredCount < totalCount) filterBadge += '<span>⚠ Filtered data (' + filteredCount + ' of ' + totalCount + ' rows)</span>';
    cftBadgeParts.forEach(p => { filterBadge += '<span>🔵 ' + p + '</span>'; });
    filterBadge += '</div>';
  }
  outDiv.innerHTML = filterBadge + '<div id="plotly-' + iid + '" style="width:100%;height:420px;"></div>';
  const plotDiv = document.getElementById('plotly-' + iid);
  // Date axis: parse + keep YYYY-MM / YYYY-MM-DD tick labels (avoid 2021-04 -> 2021)
  if (xAxisType === 'date' && layout.xaxis) {
    layout.xaxis.type = 'date';
    layout.xaxis.tickformat = _datesHaveDay(xVals) ? '%Y-%m-%d' : '%Y-%m';
    layout.xaxis.nticks = undefined;
  }
  // ── TRENDLINE / MOVING AVERAGE OVERLAY (multi-select: overlay several to compare) ──
  const _trendSel = document.getElementById('trend-' + iid);
  const _trendModes = _trendSel ? Array.from(_trendSel.selectedOptions).map(o => o.value).filter(m => m && m !== 'none') : [];
  if (_trendModes.length && !isPie && !isCombo && !isSpecial && traces.length > 0) {
    const baseT = traces[0];
    const xs = isHorizLayout ? baseT.y : baseT.x;
    const ys = (isHorizLayout ? baseT.x : baseT.y || []).map(v => parseFloat(v));
    const TPAL = ['#f7b24f','#22d3a5','#a78bfa','#f75a7a','#60a5fa','#34d399','#fb923c','#e879f9'];
    if (xs) _trendModes.forEach((mode, mi) => {
      const tlines = _computeTrend(mode, ys);
      if (!tlines) return;
      const base = TPAL[mi % TPAL.length];
      tlines.forEach(ln => {
        if (!ln || !ln.vals) return;
        const tt = { type:'scatter', mode:'lines', name: ln.name || _trendLabel(mode),
          line:{ color: base, width:2, dash: ln.dash || 'solid' }, hoverinfo:'skip' };
        if (isHorizLayout) { tt.y = xs; tt.x = ln.vals; } else { tt.x = xs; tt.y = ln.vals; }
        traces.push(tt);
      });
    });
  }

  Plotly.newPlot(plotDiv, traces, layout, config);
  // Click bar → filter table
  plotDiv.on('plotly_click', d => _onChartClick(iid, tableId, d));
  // Snapshot config so charts can be restored after a refresh
  try { _snapshotChart(iid, tableId); } catch(e) {}
}

// ── drill-down: click a chart bar/slice → filter the table to that value (and all charts) ──
// (overrides the engine's no-op stub; defined last so it wins)
function _onChartClick(iid, tableId, data){
  if(!CUR || !data || !data.points || !data.points.length) return;
  const pt=data.points[0];
  let xv = pt.x!==undefined?pt.x:(pt.label!==undefined?pt.label:null);
  if(xv===null||xv===undefined) return;
  xv=String(xv);
  const xSel=document.getElementById('xcol-'+iid); if(!xSel) return;
  const col=(_CHART_DATA&&_CHART_DATA.cols)?_CHART_DATA.cols[parseInt(xSel.value)]:null;
  if(!col) return;
  const s=st(CUR);
  const k=s.filters.findIndex(f=>f.col===col && f.op==='=' && String(f.value)===xv);
  if(k>=0){ s.filters.splice(k,1); toast('Filter cleared: '+col+' = '+xv); }
  else    { s.filters.push({col,op:'=',value:xv}); toast('Filtered: '+col+' = '+xv+' (click point again to clear)'); }
  s.page=1; renderTable(); refreshChartsForFilter();
}

initTheme();
refreshProjects();
refreshTables();
</script>
</body>
</html>
"""


_init_store()  # migrate legacy layout + open the active project (works under uvicorn too)


if __name__ == "__main__":
    try:
        import uvicorn
    except ModuleNotFoundError:
        raise SystemExit(
            'Missing uvicorn. Install: pip install "uvicorn[standard]"'
        )
    print("pudbo-polars -> http://127.0.0.1:8000")
    uvicorn.run(app, host="127.0.0.1", port=8000)
