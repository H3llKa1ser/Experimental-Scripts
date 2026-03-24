##!/usr/bin/env python3
"""
=============================================================
 TXT to CSV Parser & Data Anonymizer/Masker
 Author: Assistant
 Date: 2026
 Description:
   1. Reads raw .txt data (space/tab/custom delimited)
   2. Parses & structures it into clean .csv
   3. Masks/anonymizes sensitive fields
   4. Outputs both clean CSV and masked CSV
=============================================================

Usage:
    # Basic (space-separated txt → csv + masked csv)
    python3 txt2csv_mask.py -i data.txt

    # Custom headers
    python3 txt2csv_mask.py -i data.txt --headers "AccountID,ResourceID,Type,Access,Name,Status"

    # Tab-separated input
    python3 txt2csv_mask.py -i data.txt -d $'\t'

    # Only mask certain columns
    python3 txt2csv_mask.py -i data.txt --mask-columns 0 1 4

    # Skip masking certain columns
    python3 txt2csv_mask.py -i data.txt --skip-columns 2 3 5

    # Choose mask type
    python3 txt2csv_mask.py -i data.txt --mask-type hash

    # Custom output directory
    python3 txt2csv_mask.py -i data.txt -o /home/user/exports/

    # Multi-word column grouping
    python3 txt2csv_mask.py -i data.txt --group "3:4" --headers "AccountID,ResourceID,Type,AccessType,Name,Status"
"""

import argparse
import hashlib
import random
import string
import re
import os
import sys
import csv
from datetime import datetime


# ============================================================
# CONFIGURATION
# ============================================================

TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")


# ============================================================
# MASKING FUNCTIONS
# ============================================================

def mask_partial(value, visible_chars=4, mask_char="*"):
    """
    Keeps last N characters visible.
    123456789012 → ********9012
    """
    value = str(value).strip()
    if len(value) <= visible_chars:
        return mask_char * len(value)
    return mask_char * (len(value) - visible_chars) + value[-visible_chars:]


def mask_full(value, mask_char="*"):
    """
    Fully masks entire value.
    main-account → ************
    """
    return mask_char * len(str(value).strip())


def mask_hash(value, length=12):
    """
    SHA-256 hash (truncated).
    123456789012 → a1b2c3d4e5f6
    """
    return hashlib.sha256(str(value).encode()).hexdigest()[:length]


def mask_random_id(value, prefix="ANON_"):
    """
    Random anonymized ID.
    main-account → ANON_X7K9M2
    """
    suffix = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"{prefix}{suffix}"


def mask_email(value):
    """
    j***@***.com
    """
    value = str(value).strip()
    if "@" in value:
        local, domain = value.split("@", 1)
        ext = domain.split(".")[-1]
        return f"{local[0]}***@***.{ext}"
    return mask_partial(value)


def mask_ip(value):
    """
    192.168.xxx.xxx
    """
    value = str(value).strip()
    match = re.match(r'(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})', value)
    if match:
        return f"{match.group(1)}.{match.group(2)}.xxx.xxx"
    return mask_partial(value)


def mask_aws_account(value):
    """
    1234****9012
    """
    value = str(value).strip()
    if len(value) == 12 and value.isdigit():
        return f"{value[:4]}****{value[-4:]}"
    return mask_partial(value)


def mask_arn(value):
    """
    Masks AWS ARN, keeps service and resource type visible.
    arn:aws:iam::123456789012:role/MyRole → arn:aws:iam::1234****9012:role/****Role
    """
    value = str(value).strip()
    if value.startswith("arn:"):
        parts = value.split(":")
        # Mask account ID if present
        for i, part in enumerate(parts):
            if re.match(r'^\d{12}$', part):
                parts[i] = mask_aws_account(part)
        return ":".join(parts)
    return mask_partial(value)


# ============================================================
# AUTO-DETECTION
# ============================================================

def detect_type(value):
    """
    Auto-detects sensitive data type.
    Returns: email, ip, aws_account, arn, numeric, text
    """
    value = str(value).strip()

    if not value:
        return 'empty'

    if re.match(r'^[\w\.\+\-]+@[\w\-]+\.[a-zA-Z]{2,}$', value):
        return 'email'

    if re.match(r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$', value):
        return 'ip'

    if value.startswith("arn:"):
        return 'arn'

    if re.match(r'^\d{12}$', value):
        return 'aws_account'

    if value.isdigit():
        return 'numeric'

    return 'text'


AUTO_MASK_MAP = {
    'email':       mask_email,
    'ip':          mask_ip,
    'aws_account': mask_aws_account,
    'arn':         mask_arn,
    'numeric':     mask_partial,
    'text':        mask_partial,
    'empty':       lambda v: v,
}

MASK_STRATEGIES = {
    'partial': mask_partial,
    'full':    mask_full,
    'hash':    mask_hash,
    'random':  mask_random_id,
}


def apply_mask(value, mask_type="auto"):
    """Apply masking based on type."""
    value = str(value).strip()
    if not value:
        return value

    if mask_type == "auto":
        detected = detect_type(value)
        return AUTO_MASK_MAP[detected](value)
    elif mask_type in MASK_STRATEGIES:
        return MASK_STRATEGIES[mask_type](value)
    return mask_partial(value)


# ============================================================
# TXT PARSER
# ============================================================

def detect_delimiter(line):
    """
    Auto-detects the delimiter used in a line.
    Priority: tab → comma → pipe → multiple-spaces → single-space
    """
    if '\t' in line:
        return '\t', 'tab'
    if ',' in line:
        return ',', 'comma'
    if '|' in line:
        return '|', 'pipe'
    if '  ' in line:
        return None, 'multi-space'  # Will use regex split
    return ' ', 'space'


def smart_split(line, delimiter=None, group_indices=None):
    """
    Splits a line into fields.
    Handles multi-space, quoted fields, and grouped columns.

    group_indices: list of "start:end" strings to merge columns.
                   e.g., ["3:4"] merges columns 3 and 4 into one.
    """
    line = line.strip()

    if not line:
        return []

    # Split based on delimiter type
    if delimiter is None:
        # Multi-space: split on 2+ spaces
        fields = re.split(r'\s{2,}', line)
    elif delimiter == ' ':
        fields = line.split()
    else:
        fields = line.split(delimiter)

    # Clean whitespace
    fields = [f.strip() for f in fields]

    # Apply column grouping (merge multi-word columns)
    if group_indices:
        fields = merge_columns(fields, group_indices)

    return fields


def merge_columns(fields, group_indices):
    """
    Merges specified column ranges into single columns.
    group_indices: list of "start:end" strings, e.g., ["3:4", "6:7"]
    """
    # Parse group ranges
    groups = []
    for g in group_indices:
        start, end = map(int, g.split(":"))
        groups.append((start, end))

    # Sort groups in reverse so indices don't shift
    groups.sort(reverse=True)

    for start, end in groups:
        if start < len(fields) and end < len(fields):
            merged = " ".join(fields[start:end + 1])
            fields = fields[:start] + [merged] + fields[end + 1:]

    return fields


def parse_txt_file(input_file, delimiter=None, has_header=False,
                   headers=None, group_indices=None):
    """
    Parses a .txt file into structured rows.

    Returns:
        headers: list of column names
        rows: list of lists (each row = list of fields)
    """
    if not os.path.exists(input_file):
        print(f"[ERROR] File not found: {input_file}")
        sys.exit(1)

    with open(input_file, 'r', encoding='utf-8') as f:
        raw_lines = f.readlines()

    # Remove empty lines
    raw_lines = [line for line in raw_lines if line.strip()]

    if not raw_lines:
        print("[ERROR] File is empty.")
        sys.exit(1)

    # Auto-detect delimiter from first line
    if delimiter is None:
        delimiter, delim_name = detect_delimiter(raw_lines[0])
        print(f"[INFO] Auto-detected delimiter: {delim_name}")
    else:
        delim_name = "custom"

    # Parse all lines
    all_rows = []
    for line in raw_lines:
        fields = smart_split(line, delimiter, group_indices)
        if fields:
            all_rows.append(fields)

    if not all_rows:
        print("[ERROR] No data parsed from file.")
        sys.exit(1)

    # Determine headers
    if has_header and all_rows:
        parsed_headers = all_rows[0]
        data_rows = all_rows[1:]
    elif headers:
        parsed_headers = [h.strip() for h in headers.split(",")]
        data_rows = all_rows
    else:
        # Auto-generate column names
        max_cols = max(len(row) for row in all_rows)
        parsed_headers = [f"Column_{i+1}" for i in range(max_cols)]
        data_rows = all_rows

    # Normalize row lengths (pad short rows, trim long rows)
    num_cols = len(parsed_headers)
    normalized_rows = []
    for row in data_rows:
        if len(row) < num_cols:
            row.extend([""] * (num_cols - len(row)))
        elif len(row) > num_cols:
            row = row[:num_cols]
        normalized_rows.append(row)

    return parsed_headers, normalized_rows


# ============================================================
# CSV WRITER
# ============================================================

def write_csv(filepath, headers, rows):
    """Writes headers + rows to a CSV file."""
    with open(filepath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(headers)
        writer.writerows(rows)
    print(f"[OK] Saved: {filepath} ({len(rows)} rows)")


# ============================================================
# MASKING PROCESSOR
# ============================================================

def mask_rows(headers, rows, mask_type="auto",
              mask_columns=None, skip_columns=None):
    """
    Masks data in specified columns.

    Returns:
        masked_rows: list of masked rows
        stats: dict with masking statistics
    """
    masked_rows = []
    total_masked = 0
    type_counts = {}

    for row in rows:
        masked_row = []
        for col_idx, value in enumerate(row):
            should_mask = True

            if skip_columns and col_idx in skip_columns:
                should_mask = False

            if mask_columns is not None and col_idx not in mask_columns:
                should_mask = False

            if should_mask and value.strip():
                detected = detect_type(value)
                type_counts[detected] = type_counts.get(detected, 0) + 1
                masked_row.append(apply_mask(value, mask_type))
                total_masked += 1
            else:
                masked_row.append(value)

        masked_rows.append(masked_row)

    stats = {
        'total_rows': len(rows),
        'total_masked': total_masked,
        'type_counts': type_counts,
    }

    return masked_rows, stats


# ============================================================
# REPORT GENERATOR
# ============================================================

def print_report(input_file, csv_file, masked_file, headers,
                 stats, mask_type):
    """Prints a summary report."""
    report = f"""
╔══════════════════════════════════════════════════════════════╗
║              ✅ TXT → CSV PARSE & MASK COMPLETE             ║
╠══════════════════════════════════════════════════════════════╣
║                                                              ║
║  📥 Input file:      {input_file:<38}║
║  📄 Clean CSV:       {csv_file:<38}║
║  🔒 Masked CSV:      {masked_file:<38}║
║                                                              ║
╠══════════════════════════════════════════════════════════════╣
║  STATISTICS                                                  ║
╠══════════════════════════════════════════════════════════════╣
║  Total rows:         {stats['total_rows']:<38}║
║  Total cells masked: {stats['total_masked']:<38}║
║  Mask type:          {mask_type:<38}║
║  Columns:            {len(headers):<38}║
║                                                              ║
╠══════════════════════════════════════════════════════════════╣
║  DETECTED DATA TYPES                                         ║
╠══════════════════════════════════════════════════════════════╣"""

    for dtype, count in stats.get('type_counts', {}).items():
        report += f"\n║  {dtype:<20} {count:<39}║"

    report += f"""
║                                                              ║
╠══════════════════════════════════════════════════════════════╣
║  COLUMN HEADERS                                              ║
╠══════════════════════════════════════════════════════════════╣"""

    for i, h in enumerate(headers):
        report += f"\n║  [{i}] {h:<53}║"

    report += """
║                                                              ║
╚══════════════════════════════════════════════════════════════╝
"""
    print(report)


# ============================================================
# PREVIEW
# ============================================================

def preview_data(headers, original_rows, masked_rows, num_rows=3):
    """Shows a side-by-side preview of original vs masked data."""
    print("\n📋 PREVIEW (Original → Masked):")
    print("=" * 80)

    # Header
    print(f"{'  |  '.join(headers)}")
    print("-" * 80)

    for i in range(min(num_rows, len(original_rows))):
        orig = original_rows[i]
        mask = masked_rows[i]
        print(f"  Original: {' | '.join(orig)}")
        print(f"  Masked:   {' | '.join(mask)}")
        print(f"  {'- ' * 40}")


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="🔒 TXT → CSV Parser & Data Anonymizer",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic: auto-detect everything
  python3 txt2csv_mask.py -i data.txt

  # With custom headers
  python3 txt2csv_mask.py -i data.txt --headers "AccountID,ResourceID,Type,AccessType,AccountName,Status"

  # Group multi-word columns (merge col 3+4 into one)
  python3 txt2csv_mask.py -i data.txt --group "3:4" --headers "AccountID,ResourceID,Type,AccessType,AccountName,Status"

  # Tab-separated input
  python3 txt2csv_mask.py -i data.txt -d $'\\t'

  # Only mask columns 0 and 1
  python3 txt2csv_mask.py -i data.txt --mask-columns 0 1

  # Skip masking columns 2 and 5
  python3 txt2csv_mask.py -i data.txt --skip-columns 2 5

  # Use hash masking
  python3 txt2csv_mask.py -i data.txt --mask-type hash

  # Custom output directory
  python3 txt2csv_mask.py -i data.txt -o ./exports/

  # First line is a header
  python3 txt2csv_mask.py -i data.txt --has-header

  # Preview only (no files written)
  python3 txt2csv_mask.py -i data.txt --preview

Mask Types:
  auto     - Auto-detect & apply best mask (default)
  partial  - Show last 4 chars:    123456789012 → ********9012
  full     - Fully mask:           main-account → ************
  hash     - SHA-256 hash:         123456789012 → a1b2c3d4e5f6
  random   - Random ID:            main-account → ANON_X7K9M2

Your Data Example:
  Input:  123456789012 123456789012 Linked Role Based main-account Unknown
  Grouped (3:4): 123456789012 | 123456789012 | Linked | Role Based | main-account | Unknown
  Masked (auto): 1234****9012 | 1234****9012 | **nked | ******ased | ********ount | ***nown
        """)

    parser.add_argument('-i', '--input', required=True,
                        help='Input .txt file path')
    parser.add_argument('-o', '--output-dir', default='.',
                        help='Output directory (default: current dir)')
    parser.add_argument('-d', '--delimiter', default=None,
                        help='Input delimiter (default: auto-detect)')
    parser.add_argument('--headers', default=None,
                        help='Comma-separated column headers')
    parser.add_argument('--has-header', action='store_true',
                        help='First line of input file is a header')
    parser.add_argument('--group', nargs='+', default=None,
                        help='Merge column ranges, e.g., "3:4" merges cols 3 & 4')
    parser.add_argument('--mask-type', default='auto',
                        choices=['auto', 'partial', 'full', 'hash', 'random'],
                        help='Masking strategy (default: auto)')
    parser.add_argument('--mask-columns', nargs='+', type=int, default=None,
                        help='Only mask these column indices (0-based)')
    parser.add_argument('--skip-columns', nargs='+', type=int, default=None,
                        help='Skip masking these column indices')
    parser.add_argument('--preview', action='store_true',
                        help='Preview only, do not write files')
    parser.add_argument('--no-clean', action='store_true',
                        help='Skip writing the clean (unmasked) CSV')

    args = parser.parse_args()

    # Create output directory if needed
    os.makedirs(args.output_dir, exist_ok=True)

    # Generate output filenames
    base_name = os.path.splitext(os.path.basename(args.input))[0]
    csv_file = os.path.join(args.output_dir, f"{base_name}_parsed_{TIMESTAMP}.csv")
    masked_file = os.path.join(args.output_dir, f"{base_name}_masked_{TIMESTAMP}.csv")

    # ---- STEP 1: Parse TXT ----
    print("\n[STEP 1] Parsing TXT file...")
    headers, rows = parse_txt_file(
        input_file=args.input,
        delimiter=args.delimiter,
        has_header=args.has_header,
        headers=args.headers,
        group_indices=args.group
    )
    print(f"[OK] Parsed {len(rows)} rows × {len(headers)} columns")

    # ---- STEP 2: Mask Data ----
    print("\n[STEP 2] Masking sensitive data...")
    masked_rows, stats = mask_rows(
        headers=headers,
        rows=rows,
        mask_type=args.mask_type,
        mask_columns=args.mask_columns,
        skip_columns=args.skip_columns
    )
    print(f"[OK] Masked {stats['total_masked']} cells")

    # ---- PREVIEW ----
    preview_data(headers, rows, masked_rows)

    if args.preview:
        print("\n[INFO] Preview mode — no files written.")
        sys.exit(0)

    # ---- STEP 3: Write CSVs ----
    print("\n[STEP 3] Writing output files...")

    if not args.no_clean:
        write_csv(csv_file, headers, rows)

    write_csv(masked_file, headers, masked_rows)

    # ---- REPORT ----
    print_report(
        input_file=args.input,
        csv_file=csv_file if not args.no_clean else "SKIPPED",
        masked_file=masked_file,
        headers=headers,
        stats=stats,
        mask_type=args.mask_type
    )


if __name__ == "__main__":
    main()
