import os
import re
import sys
import json
import time
import csv
import argparse
from pathlib import Path
from typing import Any, Dict, List, Optional, Iterable

import requests
from dotenv import load_dotenv, find_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

API_PARTNUMBER = "https://api.mouser.com/api/v1.0/search/partnumber"
API_KEYWORD    = "https://api.mouser.com/api/v1.0/search/keyword"

load_dotenv(find_dotenv(), override=False)


def _get_api_key() -> str:
    key = os.getenv("MOUSER_API_KEY")
    if not key:
        print("ERROR: MOUSER_API_KEY is not set. Put it to .env or env.", file=sys.stderr)
        sys.exit(2)
    return key


_last_request_time = 0.0
_requests_today = 0

def _enforce_rate_limit():
    global _last_request_time, _requests_today
    if _requests_today >= 1000:
        print("WARNING: Reached 1000 requests limit for this session.", file=sys.stderr)
    
    now = time.time()
    elapsed = now - _last_request_time
    # 30 requests per minute -> 1 request every 2.1 seconds
    if elapsed < 2.1:
        time.sleep(2.1 - elapsed)
    _last_request_time = time.time()
    _requests_today += 1


def _post_json(url: str, params: Dict[str, str], payload: Dict[str, Any],
               max_retries: int = 3, timeout: int = 20) -> Dict[str, Any]:
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    attempt = 0
    while True:
        attempt += 1
        _enforce_rate_limit()
        try:
            resp = requests.post(url, params=params, json=payload, headers=headers, timeout=timeout)
        except requests.RequestException as e:
            if attempt <= max_retries:
                delay = 2 ** attempt
                print(f"Network error. Retrying in {delay}s...", file=sys.stderr)
                time.sleep(delay)
                continue
            raise RuntimeError(f"Network error: {e}") from e

        if resp.status_code == 429:
            if attempt <= max_retries:
                retry_after_str = resp.headers.get("Retry-After", "")
                retry_after = int(retry_after_str) if retry_after_str.isdigit() else 0
                delay = max(retry_after, 2 ** attempt)
                print(f"Rate limit hit (429). Retrying in {delay}s...", file=sys.stderr)
                time.sleep(delay)
                continue
            else:
                raise RuntimeError("Max retries exceeded for 429 Rate Limit.")

        if resp.status_code >= 400:
            raise RuntimeError(f"Mouser API error {resp.status_code}: {resp.text[:500]}")

        try:
            return resp.json()
        except json.JSONDecodeError:
            raise RuntimeError(f"Invalid JSON: {resp.text[:200]}")


def fetch_by_partnumber(query: str, api_key: str, max_retries: int, timeout: int) -> Dict[str, Any]:
    params   = {"apiKey": api_key}
    payload1 = {"SearchByPartNumberRequest": {"MouserPartNumber": query, "records": 50}}
    data = _post_json(API_PARTNUMBER, params, payload1, max_retries, timeout)
    if _has_parts(data):
        return data
    payload2 = {"SearchByPartNumberRequest": {"mouserPartNumber": query, "records": 50}}
    return _post_json(API_PARTNUMBER, params, payload2, max_retries, timeout)


def fetch_by_keyword(query: str, api_key: str, max_retries: int, timeout: int) -> Dict[str, Any]:
    params  = {"apiKey": api_key}
    payload = {"SearchByKeywordRequest": {"keyword": query, "records": 50}}
    return _post_json(API_KEYWORD, params, payload, max_retries, timeout)


def _has_parts(api_response: Dict[str, Any]) -> bool:
    sr = api_response.get("SearchResults") or api_response.get("SearchByPartNumberResponse") or {}
    parts = sr.get("Parts") or []
    return len(parts) > 0


def extract_first_part(api_response: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    sr = api_response.get("SearchResults") or api_response.get("SearchByPartNumberResponse") or {}
    parts = sr.get("Parts") or []
    return parts[0] if parts else None


_AVAIL_RE = re.compile(r"([\d,]+)")
_PN_ALLOWED_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._/-]*$")

def _is_valid_pn(q: str) -> bool:
    q = q.strip()
    if not q:
        return False
    if q in {"\\", "/"}:
        return False
    return bool(_PN_ALLOWED_RE.match(q))

def _parse_stock(availability: Optional[str]) -> Optional[int]:
    if not availability:
        return None
    m = _AVAIL_RE.search(availability)
    if m:
        try:
            return int(m.group(1).replace(",", ""))
        except ValueError:
            pass
    if "on order" in availability.lower():
        return 0
    return None

def _first_unit_price(price_breaks: Optional[List[Dict[str, Any]]]) -> Optional[str]:
    if not price_breaks:
        return None
    for pb in price_breaks:
        try:
            if int(pb.get("Quantity", 0)) == 1:
                p = pb.get("Price")
                return (str(p).strip() if p is not None else None)
        except Exception:
            continue
    p = price_breaks[0].get("Price") if price_breaks else None
    return (str(p).strip() if p is not None else None)


def transform_strict(part: Dict[str, Any]) -> Dict[str, Any]:
    category     = part.get("Category") or part.get("CategoryName") or part.get("ProductLine") or None
    availability = part.get("Availability")
    stock        = _parse_stock(availability)
    lead_time    = part.get("FactoryLeadTime") or part.get("LeadTime") or part.get("ManufacturerLeadTimeWeeks") or None
    unit_price   = _first_unit_price(part.get("PriceBreaks") or [])
    description  = part.get("Description") or part.get("ProductDescription") or None

    return {
        "Product Category": category,
        "Stock": stock,
        "Factory Lead Time": lead_time,
        "Unit Price": unit_price,
        "Description": description,
    }


def write_csv(rows: List[Dict[str, Any]], out_path: str) -> None:
    fields = list(rows[0].keys()) if rows else []
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow(r)

def print_table(rows: List[Dict[str, Any]]) -> None:
    cols = list(rows[0].keys())
    widths = {c: max(len(c), max((len(str(r.get(c) or "")) for r in rows), default=0)) for c in cols}
    header = " | ".join(c.ljust(widths[c]) for c in cols)
    sep = "-+-".join("-" * widths[c] for c in cols)
    print(header); print(sep)
    for r in rows:
        print(" | ".join(str(r.get(c) or "").ljust(widths[c]) for c in cols))


def _iter_input(parts_from_cli: List[str], input_file: Optional[str]) -> Iterable[str]:
    if input_file:
        p = Path(input_file)
        if not p.exists():
            raise FileNotFoundError(f"--input file not found: {input_file}")
        if p.suffix.lower() == ".csv":
            with p.open("r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    yield line.split(",")[0].strip()
        else:
            with p.open("r", encoding="utf-8") as f:
                for line in f:
                    s = line.strip()
                    if s:
                        yield s
    for x in parts_from_cli:
        yield x

def _to_number_or_text(value):
    if value is None:
        return None
    s = str(value).strip()
    if s.isdigit():
        try:
            return int(s)
        except Exception:
            return s
    try:
        if s.startswith("$") or s.startswith("€") or s.startswith("£"):
            s2 = s[1:]
        else:
            s2 = s
        return float(s2.replace(",", ""))
    except Exception:
        return s

def write_xlsx(rows: List[Dict[str, Any]], out_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Mouser"

    cols = list(rows[0].keys())
    header_font = Font(bold=True)
    ws.append(cols)
    for col_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for r in rows:
        out_row = []
        for c in cols:
            out_row.append(_to_number_or_text(r.get(c)))
        ws.append(out_row)

    ws.freeze_panes = "A2"

    for col_idx, col_name in enumerate(cols, start=1):
        max_len = len(str(col_name))
        for row_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None:
                continue
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    wb.save(out_path)
    print(f"Wrote XLSX -> {out_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Mouser Search CLI: Mouser PN or Manufacturer PN (keyword)."
    )
    parser.add_argument("part_numbers", nargs="*", help="PNs (Mouser or Manufacturer)")
    parser.add_argument("--input", help="Path to file with PNs (txt/csv, one PN per line or first column)")
    parser.add_argument("--format", choices=["table", "csv", "json", "xlsx"], default="table")
    parser.add_argument("--out", help="Output file path for csv/json")
    parser.add_argument("--raw", action="store_true", help="Print raw JSON to stdout")
    parser.add_argument("--save-raw", help="Directory to save raw JSON responses per PN")
    parser.add_argument("--timeout", type=int, default=20, help="HTTP timeout seconds")
    parser.add_argument("--retries", type=int, default=3, help="Max retries for 429/network errors")
    args = parser.parse_args()

    api_key = _get_api_key()
    q_iter = list(_iter_input(args.part_numbers, args.input))
    if not q_iter:
        print("Provide at least one PN via args or --input file.", file=sys.stderr)
        sys.exit(2)

    results: List[Dict[str, Any]] = []

    save_dir = None
    if args.save_raw:
        save_dir = Path(args.save_raw)
        save_dir.mkdir(parents=True, exist_ok=True)

    for q in q_iter:
        if not _is_valid_pn(q):
            print(f"WARNING: skipped invalid query token: {q!r}", file=sys.stderr)
            continue

        data_pn = fetch_by_partnumber(q, api_key, args.retries, args.timeout)
        part = extract_first_part(data_pn)
        used_keyword = False

        if not part:
            data_kw = fetch_by_keyword(q, api_key, args.retries, args.timeout)
            part = extract_first_part(data_kw)
            used_keyword = True

        if args.raw:
            if used_keyword:
                print(f"=== RAW KEYWORD RESPONSE for {q} ===")
                print(json.dumps(data_kw, ensure_ascii=False, indent=2))
            else:
                print(f"=== RAW PARTNUMBER RESPONSE for {q} ===")
                print(json.dumps(data_pn, ensure_ascii=False, indent=2))
        if save_dir is not None:
            raw_path = save_dir / f"{q.replace('/', '_')}.json"
            with raw_path.open("w", encoding="utf-8") as f:
                json.dump(data_kw if used_keyword else data_pn, f, ensure_ascii=False, indent=2)

        if not part:
            print(f"WARNING: No parts found for {q}", file=sys.stderr)
            continue

        results.append(transform_strict(part))

    if not results:
        print("No data produced.")
        return

    if args.format == "json":
        output = json.dumps(results, ensure_ascii=False, indent=2)
        if args.out:
            with open(args.out, "w", encoding="utf-8") as f:
                f.write(output)
            print(f"Wrote JSON -> {args.out}")
        else:
            print(output)
    elif args.format == "csv":
        out_path = args.out or "mouser_results.csv"
        write_csv(results, out_path)
        print(f"Wrote CSV -> {out_path}")
    elif args.format == "xlsx":
        out_path = args.out or "mouser_results.xlsx"
        write_xlsx(results, out_path)
    else:
        print_table(results)


if __name__ == "__main__":
    main()
