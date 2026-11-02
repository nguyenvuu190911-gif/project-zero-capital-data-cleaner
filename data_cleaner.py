"""Clean CSV or simple XLSX tables locally without changing the input file."""
import argparse
import csv
import json
from pathlib import Path
from xlsx_simple import read_xlsx, write_xlsx


def clean(source: Path, destination: Path, report: Path) -> dict:
    if len({p.resolve() for p in (source, destination, report)}) != 3:
        raise ValueError("Input, output and report must be separate files")
    suffix = source.suffix.lower()
    if suffix not in (".csv", ".xlsx") or destination.suffix.lower() != suffix:
        raise ValueError("Input and output must both be .csv or both be .xlsx")
    if source.stat().st_size > 20_000_000:
        raise ValueError("Input exceeds 20 MB safety limit")
    if suffix == ".csv":
        with source.open("r", encoding="utf-8-sig", newline="") as handle:
            raw = list(csv.reader(handle, strict=True))
    else:
        raw = read_xlsx(source)
    if not raw or not raw[0] or any(not str(value).strip() for value in raw[0]):
        raise ValueError("A nonempty header row is required")
    headers = [str(value).strip() for value in raw[0]]
    if len(set(headers)) != len(headers):
        raise ValueError("Headers become duplicates after trimming")
    rows, seen = [], set()
    blanks = duplicates = malformed = 0
    for row in raw[1:]:
        if len(row) > len(headers):
            malformed += 1
            continue
        values = [str(value).strip() for value in row] + [""] * (len(headers) - len(row))
        if not any(values):
            blanks += 1
            continue
        key = tuple(values)
        if key in seen:
            duplicates += 1
            continue
        seen.add(key)
        rows.append(values)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if suffix == ".csv":
        with destination.open("w", encoding="utf-8-sig", newline="") as handle:
            # Spreadsheet apps may evaluate CSV cells that begin with formula markers.
            safe = lambda value: "'" + value if value.lstrip().startswith(("=", "+", "-", "@")) else value
            csv.writer(handle).writerows([[safe(value) for value in row] for row in [headers, *rows]])
    else:
        write_xlsx(destination, [headers, *rows])
    result = {"input_rows": len(raw) - 1, "output_rows": len(rows), "blank_rows": blanks,
              "duplicate_rows": duplicates, "malformed_rows": malformed,
              "note": "XLSX output contains values only; formatting and formulas are unsupported."}
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Clean CSV or simple XLSX tables")
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    try:
        print(json.dumps(clean(args.input, args.output, args.report), ensure_ascii=False))
    except (OSError, UnicodeError, csv.Error, ValueError) as exc:
        parser.exit(2, f"Error: {exc}\n")
