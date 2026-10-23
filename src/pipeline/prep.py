"""Prep (silver layer): ingest output -> long tidy facts + a flags log.

Two structures:
- config/canonical_map.csv  how source labels, scenario words, units and metadata keys map
                            to standard names (edited by hand, no code change needed)
- facts                     one row per number; the cleaned `value` sits next to the untouched
                            `raw_value` and its `location` in the source file

Each pack goes through: pack_context -> to_long -> apply_map -> parse -> check.
Every problem found or handled is written to `flags`, using the same discrepancy_type
words as discrepancies.csv. severity "info" = handled automatically, "flag" = needs review.
"""

import csv
import re
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl.utils import get_column_letter

MAP_PATH = Path(__file__).resolve().parents[2] / "config" / "canonical_map.csv"

# metrics that are not money; everything else is converted to millions
UNITS = {"headcount": "count", "ltir": "rate", "covenant_max_leverage": "x"}
# flows add up over time (so Q2 = H1 - Q1); balances and rates do not
FLOW_METRICS = {"revenue", "ebitda", "capex", "operating_cash_flow"}
# the answer key calls the capex target "plan" and P&L targets "budget"
PLAN_METRICS = {"capex"}
RESTATED_WORDS = ("retraité", "restated")
# inputs every company must report (in answer-key terms); absent or blank -> "missing value"
REQUIRED = [
    ("revenue", "actual"), ("revenue", "budget"), ("revenue", "prior_quarter"),
    ("ebitda", "actual"), ("ebitda", "budget"), ("ebitda", "prior_quarter"),
    ("capex", "actual"), ("capex", "plan"), ("operating_cash_flow", "actual"),
    ("net_debt", "actual"), ("ebitda_ltm", "actual"), ("covenant_max_leverage", "facility"),
    ("headcount", "actual"), ("ltir", "actual"),
]

FACT_FIELDS = ["company_id", "reporting_period", "period", "metric", "scenario", "value",
               "currency", "unit", "raw_label", "raw_column", "raw_value", "source_file",
               "location"]
FLAG_FIELDS = ["company_id", "metric", "scenario", "discrepancy_type", "severity", "message",
               "source_file", "location"]

FR_MONTHS = {"janvier": 1, "février": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
             "juillet": 7, "août": 8, "septembre": 9, "octobre": 10, "novembre": 11,
             "décembre": 12}
AMOUNT = re.compile(r"[£€$]\s?(\d+(?:[.,]\d+)?)\s?(bn|m|k)?\b", re.I)


# --- small helpers ----------------------------------------------------------------------

def norm(text):
    return " ".join(str(text).split()).casefold()


def blank(value):
    return value is None or str(value).strip() == ""


def load_map(path=MAP_PATH):
    """{map_type: {normalised source_value: (canonical_value, discrepancy_type)}}"""
    cmap = defaultdict(dict)
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            cmap[row["map_type"]][norm(row["source_value"])] = (
                row["canonical_value"], row["discrepancy_type"] or None)
    return cmap


def flag(flags, ctx, dtype, message, severity="info", **where):
    flags.append({"company_id": ctx["company_id"], "discrepancy_type": dtype,
                  "severity": severity, "message": message, **where})


def emit(flags, ctx, events):
    """events: {(dtype, severity, source_file): [messages]} -> one flag row per key."""
    for (dtype, severity, source), messages in events.items():
        unique = list(dict.fromkeys(messages))
        flag(flags, ctx, dtype, "; ".join(unique), severity, source_file=source)


def quarter(d):
    return f"{d.year}-Q{(d.month - 1) // 3 + 1}"


def prev_quarter(period):
    year, q = int(period[:4]), int(period[-1])
    return f"{year - 1}-Q4" if q == 1 else f"{year}-Q{q - 1}"


def half_of(period):
    return f"{period[:4]}-H{1 if int(period[-1]) <= 2 else 2}"


def close(a, b, tol=0.005):
    if a is None or b is None:
        return a is None and b is None
    return abs(a - b) <= tol * max(abs(a), abs(b), 1e-9)


# --- parsing rules ----------------------------------------------------------------------

def parse_number(value):
    """-> (number or None, used_decimal_comma). Raises ValueError for non-numeric text."""
    if value is None:
        return None, False
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value), False
    text = str(value).strip().replace(" ", "").replace("\xa0", "").replace(" ", "")
    text = text.rstrip("xX")  # covenant multiples like "5.50x"
    if text in ("", "-", "–", "n/a"):
        return None, False
    decimal_comma = False
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):  # 1.234,5
            text, decimal_comma = text.replace(".", "").replace(",", "."), True
        else:  # 1,234.5
            text = text.replace(",", "")
    elif "," in text:
        if re.fullmatch(r"-?\d{1,3}(,\d{3})+", text):  # 1,600 = thousands separator
            text = text.replace(",", "")
        else:  # 61800,0 / 27,4
            text, decimal_comma = text.replace(",", "."), True
    return float(text), decimal_comma


def parse_date(value, date_format=None):
    """-> (date, how) where how is iso / excel serial / day-first / month-first / french text."""
    if isinstance(value, datetime):
        return value.date(), "iso"
    if isinstance(value, (int, float)) and 20000 < value < 80000:
        return date(1899, 12, 30) + timedelta(days=int(value)), "excel serial"
    text = str(value or "")
    if iso := re.findall(r"(\d{4})-(\d{2})-(\d{2})", text):
        y, m, d = map(int, iso[-1])  # last date wins: "2026-04-01 to 2026-06-30"
        return date(y, m, d), "iso"
    if m := re.search(r"(\d{1,2})[/.](\d{1,2})[/.](\d{4})", text):
        a, b, y = map(int, m.groups())
        if str(date_format or "").upper().startswith("MM"):
            return date(y, a, b), "month-first"
        return date(y, b, a), "day-first"
    if m := re.search(r"(\d{1,2})\s+([a-zéû]+)\s+(\d{4})", text.lower()):
        if m[2] in FR_MONTHS:
            return date(int(m[3]), FR_MONTHS[m[2]], int(m[1])), "french text"
    return None, None


def parse_period(text, ctx):
    """'Q2 2026', 'T1 2026', 'Q2-2026', '2026-Q2', 'Q1 FY27', 'H1 2026', 'as at 06/30/2026'
    -> ('2026-Q2' / '2026-H1' / None, used_fiscal_calendar)"""
    if blank(text):
        return None, False
    text = str(text)
    if m := re.search(r"\b[QT]([1-4])\s*FY\s*(\d{2,4})", text, re.I):
        n, fy, start = int(m[1]), 2000 + int(m[2]) % 100, ctx["fiscal_start"]
        month, year = start + 3 * (n - 1), fy - 1 if start > 1 else fy
        if month > 12:
            month, year = month - 12, year + 1
        return quarter(date(year, month, 1)), True
    if m := re.search(r"\b[QT]([1-4])[\s-]*(\d{4})", text):
        return f"{m[2]}-Q{m[1]}", False
    if m := re.search(r"(\d{4})-?Q([1-4])", text):
        return f"{m[1]}-Q{m[2]}", False
    if m := re.search(r"\bH([12])\s*(\d{4})", text):
        return f"{m[2]}-H{m[1]}", False
    d, _ = parse_date(text, ctx["date_format"])
    return (quarter(d) if d else None), False


# --- step 0: pack context (metadata) ------------------------------------------------------

def pack_context(company_id, files, cmap, flags):
    ctx = {"company_id": company_id, "currency": None, "unit": None, "date_format": None,
           "fiscal_start": 1, "period_end": None, "texts": []}
    meta_files = [f for f in files if f["file"].startswith("metadata.")]
    source = meta_files[0]["file"] if meta_files else ""
    meta, renamed = {}, []
    for f in meta_files:
        for key, value in (f["data"] or {}).items():
            hit = cmap["meta_key"].get(norm(key))
            if hit is None:
                continue
            meta[hit[0]] = value
            if hit[0] != key:
                renamed.append(f"{key} -> {hit[0]}")
    if renamed:
        flag(flags, ctx, "different metadata keys", "; ".join(renamed), source_file=source)

    ctx.update(currency=meta.get("currency"), unit=meta.get("unit"),
               date_format=meta.get("date_format"),
               fiscal_start=int(meta.get("fiscal_year_start_month") or 1))
    for key in ("period_end", "submitted"):
        if key in meta:
            d, how = parse_date(meta[key], ctx["date_format"])
            ctx[key] = d
            if how not in (None, "iso"):
                flag(flags, ctx, "date format", f"{key} {meta[key]!r} read {how} as {d}",
                     source_file=source)
    ctx["reporting_period"] = (quarter(ctx["period_end"]) if ctx["period_end"]
                               else parse_period(meta.get("period_label"), ctx)[0])
    if ctx["currency"] and ctx["currency"] != "GBP":
        flag(flags, ctx, "currency",
             f"values kept in {ctx['currency']}; GBP conversion happens in gold",
             source_file=source)
    return ctx


# --- step 1: to_long (every table -> rows of label / column / value) ----------------------

def _row(f, kind, location, label, column, value, units=(), titles=(), hidden=False):
    return {"source_file": f["file"], "source_type": kind, "location": location,
            "raw_label": label, "raw_column": column, "raw_value": value,
            "unit_texts": [u for u in units if not blank(u)], "title_texts": list(titles),
            "hidden": hidden}


def _excel_rows(f):
    """Split each sheet into blocks at blank rows. One-cell blocks are titles/notes;
    the rest are tables whose all-text rows are headers (merged headers filled right)."""
    out, titles_all = [], []
    for t in f["tables"]:
        grid, hidden = t["rows"], set(t["hidden_rows"])
        blocks, current = [], []
        for i, row in enumerate(grid):
            if all(blank(v) for v in row):
                if current:
                    blocks.append(current)
                current = []
            else:
                current.append(i)
        if current:
            blocks.append(current)
        tables = [b for b in blocks if max(sum(not blank(v) for v in grid[i]) for i in b) >= 2]
        titles = [str(v) for b in blocks if b not in tables for i in b for v in grid[i]
                  if not blank(v)]
        titles_all += titles

        for block in tables:
            width = max(len(grid[i]) for i in block)
            rows = {i: list(grid[i]) + [None] * (width - len(grid[i])) for i in block}
            header_idx = [i for k, i in enumerate(block) if _is_header(rows[i], k == 0)]
            for n, i in enumerate(header_idx):  # merged cell: repeat across while a sub-header exists
                below = rows[header_idx[n + 1]] if n + 1 < len(header_idx) else None
                for c in range(2, width):
                    if blank(rows[i][c]) and below and not blank(below[c]):
                        rows[i][c] = rows[i][c - 1]
            headers = [" | ".join(dict.fromkeys(str(rows[i][c]) for i in header_idx
                                                if not blank(rows[i][c]))) or None
                       for c in range(width)]
            for i in block:
                if i in header_idx:
                    continue
                for c in range(1, width):
                    if headers[c] is None and blank(rows[i][c]):
                        continue
                    out.append(_row(f, "excel", f"{t['name']}!{get_column_letter(c + 1)}{i + 1}",
                                    rows[i][0], headers[c], rows[i][c],
                                    units=[headers[c], headers[0]], titles=titles,
                                    hidden=(i + 1) in hidden))
    return out, titles_all


def _is_header(row, first_in_block):
    rest = [v for v in row[1:] if not blank(v)]
    return bool(rest) and all(isinstance(v, str) for v in rest) and (blank(row[0]) or first_in_block)


def _grid_rows(f, kind):
    """csv / pdf / docx tables: first row is the header, first column the label.
    A header with a 'value' column means the table is already long (A2): the other
    columns then describe each value instead of being values themselves."""
    out = []
    for t in f["tables"]:
        header, prefix = t["rows"][0], "" if kind == "csv" else f"{t['name']}."
        value_cols = [c for c, h in enumerate(header) if h and "value" in str(h).lower()]
        for r, row in enumerate(t["rows"][1:], start=2):
            if value_cols:
                described_by = " | ".join(str(row[c]) for c in range(1, len(row))
                                          if c not in value_cols and not blank(row[c]))
                for c in value_cols:
                    out.append(_row(f, kind, f"{prefix}r{r}c{c + 1}", row[0], described_by,
                                    row[c], units=[header[c]]))
            else:
                for c in range(1, len(row)):
                    out.append(_row(f, kind, f"{prefix}r{r}c{c + 1}", row[0], header[c],
                                    row[c], units=[header[c], header[0]]))
    return out


def _json_rows(f, cmap):
    """Lists of records: first key is the label, keys that are scenario words hold values,
    the remaining keys (e.g. period) describe them. Top-level strings are unit hints."""
    out, data = [], f["data"]
    if not isinstance(data, dict):
        return out
    hints = [v for v in data.values() if isinstance(v, str)]
    for key, records in data.items():
        if not (isinstance(records, list) and records and all(isinstance(r, dict) for r in records)):
            continue
        label_key = next(iter(records[0]))
        for i, rec in enumerate(records):
            value_keys = [k for k in rec if norm(k) in cmap["scenario"]]
            described_by = " | ".join(str(v) for k, v in rec.items()
                                      if k != label_key and k not in value_keys and v is not None)
            for k in value_keys:
                out.append(_row(f, "json", f"{key}[{i}].{k}", rec[label_key],
                                f"{described_by} | {k}", rec[k], units=hints))
    return out


def _find_metric_word(sentence, cmap):
    found = [(m.start(), -len(word), sentence[m.start():m.end()])
             for word in cmap["metric"]
             if (m := re.search(rf"(?<!\w){re.escape(word)}(?!\w)", sentence.casefold()))]
    return min(found) if found else None


def _text_rows(f, cmap):
    """Prose source (A4 CFO email): '<metric word> ... £88.6m ... plan £90.0m'.
    First amount after the metric word is the actual; a second one is the plan."""
    out = []
    for p, para in enumerate(f["text"], start=1):
        for s, sentence in enumerate(re.split(r"(?<=[.;!?])\s+", para), start=1):
            hit = _find_metric_word(sentence, cmap)
            if hit is None:
                continue
            start, _, word = hit
            after = sentence[start:]
            for n, (number, scale) in enumerate(AMOUNT.findall(after)[:2]):
                if n == 1 and "plan" not in after.lower():
                    break
                unit = {"bn": "bn", "k": "thousands"}.get(scale.lower(), "millions")
                out.append(_row(f, "text", f"para{p}.s{s}", word, "plan" if n else "actual",
                                number, units=[unit]))
    return out


def to_long(files, ctx, cmap):
    rows = []
    for f in files:
        if f["file"].startswith("metadata."):
            continue
        ctx["texts"] += f["text"]
        if f["type"] == "excel":
            excel_rows, titles = _excel_rows(f)
            rows += excel_rows
            ctx["texts"] += titles
        elif f["type"] in ("csv", "pdf", "docx"):
            rows += _grid_rows(f, f["type"])
        elif f["type"] in ("json", "yaml"):
            rows += _json_rows(f, cmap)
        elif f["type"] == "txt":
            rows += _text_rows(f, cmap)
    return rows


# --- step 2: apply_map (canonical names) -------------------------------------------------

def _scenario(text, cmap):
    hits = [(m.start(), canon) for word, (canon, _) in cmap["scenario"].items()
            if (m := re.search(rf"(?<!\w){re.escape(word)}(?!\w)", norm(text or "")))]
    return min(hits)[1] if hits else "actual"


def _is_number(value):
    try:
        return parse_number(value)[0] is not None
    except ValueError:
        return False


def apply_map(rows, cmap, ctx, flags):
    mapped, events = [], defaultdict(list)
    for row in rows:
        label, column = row["raw_label"], row["raw_column"]
        if blank(label):
            continue
        hit, period_text = cmap["metric"].get(norm(label)), column
        if hit is None and column and norm(column) in cmap["metric"] and parse_period(label, ctx)[0]:
            # transposed table: periods down the side, metric across the top (A4 Safety tab)
            hit, period_text, column = cmap["metric"][norm(column)], label, None
        if hit is None:
            if _is_number(row["raw_value"]):
                events[("unmapped label", "flag", row["source_file"])].append(
                    f"{label!r} not in canonical_map.csv")
            continue  # text-only rows (risk registers etc.) are not facts
        metric, dtype = hit
        if dtype:
            events[(dtype, "info", row["source_file"])].append(f"{label!r} -> {metric}")
        mapped.append({**row, "metric": metric, "period_text": period_text,
                       "scenario": _scenario(column, cmap)})
    emit(flags, ctx, events)
    return mapped


# --- step 3: parse (numbers, units, periods) ---------------------------------------------

def _unit_scale(row, ctx, cmap):
    """-> (scale, came_from_title_row). Looks in headers first, then title rows, then metadata."""
    for texts, from_title in ((row["unit_texts"], False), (row["title_texts"], True),
                              ([ctx["unit"]], False)):
        for text in texts:
            for token, (scale, _) in cmap["unit"].items():
                if token in norm(text or ""):
                    return float(scale), from_title
    return 1.0, False


def parse(rows, ctx, cmap, flags):
    facts, events = [], defaultdict(list)
    for row in rows:
        try:
            value, decimal_comma = parse_number(row["raw_value"])
        except ValueError:
            continue  # text in a numeric table, e.g. "Aggravé"
        src = row["source_file"]
        if row["metric"] == "period_end":
            d, how = parse_date(row["raw_value"])
            if how == "excel serial":
                events[("dates", "info", src)].append(f"Excel serial {row['raw_value']} -> {d}")
            continue

        period, fiscal = parse_period(row["period_text"], ctx)
        period = period or ctx["reporting_period"]
        if fiscal:
            events[("period definition", "info", src)].append(
                f"{row['period_text']!r} -> {period} (fiscal year starts month {ctx['fiscal_start']})")

        metric, scenario = row["metric"], row["scenario"]
        if metric == "covenant_max_leverage" and scenario == "actual":
            scenario = "facility"  # the covenant is a facility term, not an actual result
        elif scenario == "budget" and metric in PLAN_METRICS:
            scenario = "plan"
        elif scenario == "plan" and metric not in PLAN_METRICS:
            scenario = "budget"

        unit = UNITS.get(metric, "m")
        if unit == "m" and value is not None:
            scale, from_title = _unit_scale(row, ctx, cmap)
            if scale != 1:
                events[("units", "info", src)].append(f"scaled x{scale:g} to millions")
            if from_title:
                events[("semi-structured layout", "info", src)].append("units read from title row")
            value = round(value * scale, 6)
        if decimal_comma:
            events[("format", "info", src)].append("decimal comma")
        if row["hidden"]:
            events[("semi-structured layout", "info", src)].append(
                f"hidden row read: {row['raw_label']}")

        facts.append({
            "company_id": ctx["company_id"], "reporting_period": ctx["reporting_period"],
            "period": period, "metric": metric, "scenario": scenario, "value": value,
            "currency": ctx["currency"], "unit": unit, "raw_label": row["raw_label"],
            "raw_column": row["raw_column"],
            "raw_value": None if row["raw_value"] is None else str(row["raw_value"]),
            "source_file": src, "location": row["location"], "source_type": row["source_type"],
        })
    emit(flags, ctx, events)
    return facts


# --- step 4: check (derive, de-duplicate, quality checks) --------------------------------

def _derive_quarter_from_half(facts, ctx, flags):
    rp = ctx["reporting_period"]
    if not rp.endswith("Q2"):
        return facts
    h1, q1 = f"{rp[:4]}-H1", prev_quarter(rp)
    have = {(f["metric"], f["scenario"], f["period"]): f for f in facts}
    derived, rate_only = [], set()
    for (metric, scenario, period), half in list(have.items()):
        if period != h1:
            continue
        if metric not in FLOW_METRICS:
            if (metric, scenario, rp) not in have:
                rate_only.add(metric)
            continue
        first = have.get((metric, scenario, q1))
        if (metric, scenario, rp) in have or first is None or None in (half["value"], first["value"]):
            continue
        derived.append({**half, "period": rp, "value": round(half["value"] - first["value"], 6),
                        "raw_value": f"{half['raw_value']} - {first['raw_value']}",
                        "location": f"{half['location']} - {first['location']}"})
    if derived:
        names = sorted({d["metric"] for d in derived})
        flag(flags, ctx, "period definition", f"Q2 derived as H1 YTD - Q1 for {', '.join(names)}",
             source_file=derived[0]["source_file"])
    for metric in sorted(rate_only):
        flag(flags, ctx, "period definition", f"{metric} reported for H1 YTD only, not Q2",
             "flag", metric=metric)
    return facts + derived


def _priority(f):
    text = f"{f['raw_column']} {f['source_file']}".casefold()
    restated = any(w in text for w in RESTATED_WORDS) - ("original" in f["source_file"].casefold())
    return (restated, f["value"] is not None, f["source_type"] != "text")


def _resolve_duplicates(facts, ctx, flags):
    """Same metric/scenario/period from several places: keep one, flag what differed."""
    groups = defaultdict(list)
    for f in facts:
        groups[(f["metric"], f["scenario"], f["period"])].append(f)
    kept, prose = [], []
    for (metric, scenario, period), group in groups.items():
        best = max(group, key=_priority)
        kept.append(best)
        for other in (g for g in group if g is not best):
            same = close(other["value"], best["value"])
            if "text" in (other["source_type"], best["source_type"]):
                prose.append((metric, scenario, other, best, same))
            elif same:
                continue
            elif _priority(best)[0] > _priority(other)[0]:
                flag(flags, ctx, "restatement",
                     f"{period} {metric}: {other['value']} in {other['source_file']} restated to "
                     f"{best['value']}; using restated", "flag", metric=metric,
                     scenario=scenario, source_file=best["source_file"], location=best["location"])
            else:
                flag(flags, ctx, "source conflict",
                     f"{period} {metric}: {other['value']} ({other['source_file']}) vs "
                     f"{best['value']} ({best['source_file']}); using {best['source_file']}",
                     "flag", metric=metric, scenario=scenario)
    if prose:
        text_file = next(o["source_file"] if o["source_type"] == "text" else b["source_file"]
                         for _, _, o, b, _ in prose)
        bad = [f"{m}/{s}: {o['value']} vs {b['value']}" for m, s, o, b, same in prose if not same]
        message = (f"{len(prose)} values read from prose; "
                   + (f"{len(bad)} disagree: {'; '.join(bad)}" if bad else "all agree with tables"))
        flag(flags, ctx, "unstructured source", message, "flag" if bad else "info",
             source_file=text_file)
    return kept


def _explanation(ctx, metric, cmap):
    words = [w for w, (canon, _) in cmap["metric"].items() if canon == metric]
    for para in ctx["texts"]:
        for sentence in re.split(r"(?<=[.;!?])\s+", str(para)):
            low = sentence.casefold()
            if any(w in low for w in words) and re.search(r"because|due to|driven by|brought forward", low):
                return sentence.strip()
    return None


def check(facts, ctx, cmap, flags):
    facts = _derive_quarter_from_half(facts, ctx, flags)
    facts = _resolve_duplicates(facts, ctx, flags)
    rp = ctx["reporting_period"]
    now = {(f["metric"], f["scenario"]): f for f in facts if f["period"] == rp}

    for metric in sorted(FLOW_METRICS):
        actual = now.get((metric, "actual"))
        target = now.get((metric, "plan" if metric in PLAN_METRICS else "budget"))
        if not (actual and target and actual["value"] and target["value"]):
            continue
        ratio = actual["value"] / target["value"]
        if 8 <= ratio <= 12:  # ~10x the target: almost always a misplaced decimal
            fixed = round(actual["value"] / 10, 6)
            flag(flags, ctx, "outlier / unit error",
                 f"{metric} {actual['value']} is {ratio:.1f}x target {target['value']}; "
                 f"using {fixed} (raw value kept)", "flag", metric=metric, scenario="actual",
                 source_file=actual["source_file"], location=actual["location"])
            actual["value"] = fixed
            ratio = fixed / target["value"]
        if abs(ratio - 1) > 0.25:
            why = _explanation(ctx, metric, cmap)
            flag(flags, ctx, "variance",
                 f"{metric} {ratio - 1:+.0%} vs {target['scenario']}"
                 + (f" - explained: {why!r}" if why else " - no explanation found"),
                 "flag", metric=metric, scenario="actual", source_file=actual["source_file"],
                 location=actual["location"])

    for (metric, scenario), current in now.items():
        history = [f["value"] for f in facts if f["metric"] == metric and f["scenario"] == scenario
                   and f["period"] < rp and "-Q" in f["period"] and f["value"] is not None]
        if len(history) >= 3 and current["value"] is not None and max(history) > 0:
            if current["value"] > 3 * max(history):
                flag(flags, ctx, "anomaly",
                     f"{metric} {current['value']} vs history {min(history)}-{max(history)}",
                     "flag", metric=metric, scenario=scenario,
                     source_file=current["source_file"], location=current["location"])

    view = kpi_inputs(facts)
    for metric, scenario in REQUIRED:
        found = view.get((ctx["company_id"], metric, scenario))
        if found is None or found["value"] is None:
            flag(flags, ctx, "missing value", f"{metric} {scenario} not reported; left empty",
                 "flag", metric=metric, scenario=scenario,
                 source_file=found["source_file"] if found else "",
                 location=found["location"] if found else "")
    return facts


# --- answer-key view + entry point --------------------------------------------------------

def kpi_inputs(facts):
    """{(company_id, metric, scenario): fact} in answer-key terms: reporting-quarter values,
    previous-quarter actuals as 'prior_quarter', and H1-only rates (flagged in check)."""
    out, fallback = {}, {}
    for f in facts:
        rp = f["reporting_period"]
        if f["period"] == rp:
            out.setdefault((f["company_id"], f["metric"], f["scenario"]), f)
        elif f["period"] == prev_quarter(rp) and f["scenario"] == "actual":
            out.setdefault((f["company_id"], f["metric"], "prior_quarter"), f)
        elif f["period"] == half_of(rp) and f["metric"] not in FLOW_METRICS:
            fallback.setdefault((f["company_id"], f["metric"], f["scenario"]), f)
    return {**fallback, **out}


def prep_set(ingested, canonical_map=None):
    """ingest_set() output -> (facts, flags), both lists of dicts."""
    cmap = canonical_map or load_map()
    facts, flags = [], []
    for pack, files in ingested.items():
        ctx = pack_context(pack.split("_", 1)[0], files, cmap, flags)
        rows = apply_map(to_long(files, ctx, cmap), cmap, ctx, flags)
        facts += check(parse(rows, ctx, cmap, flags), ctx, cmap, flags)
    return facts, flags


def write_csv(path, rows, fields):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore", restval="")
        writer.writeheader()
        writer.writerows(rows)
