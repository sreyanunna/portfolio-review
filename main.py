"""Run the pipeline: ingest (bronze) -> prep (silver).

    uv run main.py                                  # Set A
    uv run main.py data/set_B_asset_managers

Writes output/silver_facts.csv and output/silver_flags.csv.
"""

import sys
from collections import Counter

from pipeline.ingest import ingest_set
from pipeline.prep import FACT_FIELDS, FLAG_FIELDS, prep_set, write_csv


def main():
    set_dir = sys.argv[1] if len(sys.argv) > 1 else "data/set_A_infrastructure"
    ingested = ingest_set(set_dir)
    facts, flags = prep_set(ingested)

    write_csv("output/silver_facts.csv", facts, FACT_FIELDS)
    write_csv("output/silver_flags.csv", flags, FLAG_FIELDS)

    fact_counts = Counter(f["company_id"] for f in facts)
    flag_counts = Counter((f["company_id"], f["severity"]) for f in flags)
    print(f"{'company':<8} {'facts':>5} {'info':>5} {'flags':>5}")
    for company in sorted(fact_counts):
        print(f"{company:<8} {fact_counts[company]:>5} {flag_counts[(company, 'info')]:>5} "
              f"{flag_counts[(company, 'flag')]:>5}")
    print("\n-> output/silver_facts.csv, output/silver_flags.csv")


if __name__ == "__main__":
    main()
