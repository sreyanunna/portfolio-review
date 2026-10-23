"""Silver layer checks for Set A: values vs answer key, flags vs discrepancies, raw kept."""

import csv
from datetime import date
from pathlib import Path

import pytest

from pipeline.ingest import ingest_set
from pipeline.prep import kpi_inputs, parse_date, parse_number, parse_period, prep_set

SET_A = Path(__file__).resolve().parents[1] / "data" / "set_A_infrastructure"
INPUT_METRICS = {"revenue", "ebitda", "capex", "operating_cash_flow", "net_debt", "ebitda_ltm",
                 "covenant_max_leverage", "headcount", "ltir"}
# need calculated KPIs, so they are checked in the gold layer, not here
GOLD_TYPES = {"threshold breach", "threshold near-breach", "narrative contradicts numbers"}


def read_csv(path):
    with open(path, encoding="utf-8") as f:
        return list(csv.DictReader(f))


@pytest.fixture(scope="module")
def ingested():
    return ingest_set(SET_A)


@pytest.fixture(scope="module")
def silver(ingested):
    return prep_set(ingested)


# --- parsing rules -----------------------------------------------------------------------

@pytest.mark.parametrize(("raw", "expected"), [
    ("61800,0", 61800.0), ("27,4", 27.4), ("5.50x", 5.5), ("7,00x", 7.0), ("1,600", 1600.0),
    ("15 600", 15600.0), (42.1, 42.1), ("-", None), (None, None),
])
def test_parse_number(raw, expected):
    assert parse_number(raw)[0] == expected


def test_parse_number_rejects_text():
    with pytest.raises(ValueError):
        parse_number("Aggravé")


@pytest.mark.parametrize(("raw", "fmt", "expected"), [
    ("30/06/2026", None, date(2026, 6, 30)), ("06/30/2026", "MM/DD/YYYY", date(2026, 6, 30)),
    ("24.07.2026", "DD.MM.YYYY", date(2026, 7, 24)), ("30 juin 2026", None, date(2026, 6, 30)),
    (46203, None, date(2026, 6, 30)), ("2026-04-01 to 2026-06-30", None, date(2026, 6, 30)),
])
def test_parse_date(raw, fmt, expected):
    assert parse_date(raw, fmt)[0] == expected


@pytest.mark.parametrize(("text", "fiscal_start", "expected"), [
    ("Q2 2026 (Apr-Jun) | Actual", 1, "2026-Q2"), ("T1 2026 Réel (retraité)", 1, "2026-Q1"),
    ("Q2-2026 | ACT", 1, "2026-Q2"), ("Q1 FY27 Actual", 4, "2026-Q2"),
    ("Q4 FY26 Actual", 4, "2026-Q1"), ("H1 2026 YTD | actual", 1, "2026-H1"),
])
def test_parse_period(text, fiscal_start, expected):
    ctx = {"fiscal_start": fiscal_start, "date_format": None}
    assert parse_period(text, ctx)[0] == expected


# --- silver vs the test oracle -------------------------------------------------------------

def test_inputs_match_answer_key(silver):
    facts, _ = silver
    got = kpi_inputs(facts)
    rows = [r for r in read_csv(SET_A / "answer_key.csv") if r["metric"] in INPUT_METRICS]
    misses = []
    for r in rows:
        fact = got.get((r["company_id"], r["metric"], r["scenario"]))
        value = fact["value"] if fact else None
        if r["expected_value_native"] == "":
            ok = value is None
        else:
            expected = float(r["expected_value_native"])
            ok = value is not None and abs(value - expected) <= 0.005 * abs(expected)
        if not ok:
            misses.append(f"{r['company_id']} {r['metric']} {r['scenario']}: "
                          f"expected {r['expected_value_native'] or 'empty'}, got {value}")
    print(f"\nanswer key inputs: {len(rows) - len(misses)}/{len(rows)} correct")
    assert misses == []


def test_every_silver_discrepancy_is_flagged(silver):
    _, flags = silver
    raised = {(f["company_id"], f["discrepancy_type"]) for f in flags}
    missing = []
    for r in read_csv(SET_A / "discrepancies.csv"):
        if r["discrepancy_type"] in GOLD_TYPES:
            continue
        for part in r["discrepancy_type"].split(" + "):  # e.g. "units + dates"
            if (r["company_id"], part) not in raised:
                missing.append(f"{r['company_id']} {part}: {r['what_was_planted']}")
    assert missing == []


def test_raw_values_are_kept_unchanged(ingested, silver):
    facts, _ = silver
    blobs = {(pack.split("_")[0], f["file"]): repr(f) for pack, files in ingested.items()
             for f in files}
    for fact in facts:
        assert fact["source_file"] and fact["location"]
        if fact["raw_value"] is None:
            continue
        blob = blobs[(fact["company_id"], fact["source_file"])]
        for part in fact["raw_value"].split(" - "):  # derived rows show both inputs
            assert part in blob, (fact["location"], part)
