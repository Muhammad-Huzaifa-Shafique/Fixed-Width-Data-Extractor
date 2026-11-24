#https://chatgpt.com/share/6930b733-6028-8005-9bab-fa8aa6303f0d
# I worked with Adil we worked together on this project and saw eachothers output and we both got the same output.

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auto fixed-width TXT → polished Excel converter.

This script:
  * Opens a panel (file dialog) so you can choose a TXT file.
  * Detects column boundaries from the first 3 header lines:
        - columns are separated by runs of 4+ spaces
        - long column names may wrap vertically across the 3 lines
  * Detects data rows automatically:
        - a new logical row starts when column 19 (index 18) is non-space
        - lines with blank column 19 are continuation lines
  * Reconstructs wrapped field values across multiple physical lines:
        - if previous segment filled the full column width → join with NO space
        - otherwise → join with a single space
  * Writes a clean .xlsx file in the same folder as the TXT file,
    with the same base name + "_clean.xlsx".
  * Then “polishes” the Excel:
        - trims leading/trailing spaces in every cell
        - bold header row
        - wraps text
        - auto-sizes columns (reasonable width limit)
        - freezes header row and adds an auto-filter

Third-party packages:
    pandas   – DataFrame handling and Excel export
    openpyxl – Excel writer engine and styling

Install if needed:
    pip install pandas openpyxl
"""

import os
import re
import sys

try:
    import pandas as pd
except ImportError:
    print("ERROR: This script requires the 'pandas' package (and 'openpyxl' for Excel writing).")
    print("Install them with:\n  pip install pandas openpyxl")
    sys.exit(1)

# Column index (0-based) that marks the start of a new logical record
ROW_START_COL = 18  # column 19 in human terms


def ask_for_file() -> str:
    """
    Ask the user to select the TXT file.

    1. Try to show a graphical file-selection panel (tkinter.filedialog).
    2. If the user cancels OR tkinter is not available, fall back to input().

    Returns
    -------
    str
        Absolute path to the selected TXT file.
    """
    path = None

    # --- Try GUI file dialog first ---
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()  # Hide the root window (we only want the dialog).

        print("A file-selection panel should be open now. Choose your fixed-width TXT file.")
        path = filedialog.askopenfilename(
            title="Select fixed-width TXT file",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )

        root.destroy()
    except Exception:
        # On systems without tkinter or if something else goes wrong,
        # we'll fall back to console input.
        path = None

    # --- Fallback to manual path entry if needed ---
    if not path:
        path = input("Enter the path to the fixed-width TXT file: ").strip().strip('"')

    if not path:
        print("No file selected. Exiting.")
        sys.exit(1)

    if not os.path.isfile(path):
        print(f"Cannot find file: {path}")
        sys.exit(1)

    return os.path.abspath(path)


def detect_column_spans(header_lines):
    """
    Detect fixed-width column spans from the first three header lines.

    HOW COLUMN WIDTHS ARE DETECTED
    ------------------------------
    * Rule: columns in this file are separated by **four or more spaces**.
    * We:
        1. Choose the header line that has the most 4+ space gaps.
        2. In that line, every run of 4+ spaces marks a column boundary.
        3. The text regions between those gaps become column slices.

    Parameters
    ----------
    header_lines : list[str]
        First three lines of the file.

    Returns
    -------
    boundary_line : str
        The header line we used to determine the spans.
    spans : list[tuple[int, int]]
        List of (start, end) indices for each column (Python slice style).
    """

    def count_gaps(line: str) -> int:
        return len(list(re.finditer(r" {4,}", line)))

    # Pick the header line with the most gaps as the "layout template".
    boundary_index = max(range(len(header_lines)), key=lambda i: count_gaps(header_lines[i]))
    boundary_line = header_lines[boundary_index]

    # Find all 4+ space runs on that line.
    gaps = list(re.finditer(r" {4,}", boundary_line))

    # Column starts: first column starts at 0, others start after a gap ends.
    starts = [0] + [g.end() for g in gaps]
    # Column ends: last column ends at line end, others end where gaps start.
    ends = [g.start() for g in gaps] + [len(boundary_line)]

    spans = list(zip(starts, ends))
    return boundary_line, spans


def build_column_names(header_lines, spans):
    """
    Merge the 3 header lines into single-line column names.

    HOW HEADER ROWS ARE MERGED
    --------------------------
    For each column span:
      * Take that character slice from each header line.
      * If the slice has non-space characters, keep it as a "piece".
      * Join pieces from top to bottom with spaces.
      * Clean filler characters (dots, &, extra whitespace).

    Example:
        Line1: 'Field......'
        Line2: 'Name.......'
        → 'Field Name'

    Parameters
    ----------
    header_lines : list[str]
    spans : list[tuple[int, int]]

    Returns
    -------
    list[str]
        Clean column names, one per span.
    """
    names = []

    for start, end in spans:
        pieces = []
        for h in header_lines:
            seg = h[start:end]
            if seg.strip():
                pieces.append(seg.strip())

        raw_name = " ".join(pieces)

        # Clean filler dots & symbols, collapse spaces.
        cleaned = raw_name.replace(".", " ").replace("&", " ")
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        if not cleaned:
            cleaned = f"Column_{len(names) + 1}"

        names.append(cleaned)

    return names


def pad_to_width(line: str, target_len: int) -> str:
    """
    Right-pad or truncate a line to `target_len` characters.

    Fixed-width parsing assumes that every column slice [start:end]
    exists; padding ensures we don't get IndexErrors when a physical
    line is shorter than the header line.
    """
    if len(line) < target_len:
        return line.ljust(target_len)
    return line[:target_len]


def split_line_into_segments(line: str, spans, target_len: int):
    """
    Split one physical line into raw column segments.

    We **keep trailing spaces** in each segment, because they tell us
    whether the text filled the entire column (important when deciding
    whether to join a continuation with or without a space).
    """
    line = pad_to_width(line, target_len)
    return [line[start:end] for (start, end) in spans]


def parse_records(body_lines, boundary_line: str, spans):
    """
    Parse the data part of the file into logical records.

    HOW CONTINUATION LINES ARE IDENTIFIED
    -------------------------------------
    * After the header:
        - A new record starts when column 19 (index 18) is non-space.
        - If column 19 is space/blank, the line continues the previous record.

    HOW WRAPPED TEXT IS RECONSTRUCTED
    ---------------------------------
    For each column, we keep track of the last raw padded segment
    (with spaces). When a continuation line adds more text in the
    same column, we:

        let width = len(last_raw_segment)
        let prev_trim = last_raw_segment.rstrip()

        - If len(prev_trim) == width:
              previous segment reached the column edge
              → treat as mid-word cut, append with **no extra space**.
        - Else:
              previous segment ended early (internal spaces at the end)
              → append with **one space**.

    This rule follows your specification for joining wrapped values.
    """
    target_len = len(boundary_line)

    def is_new_record(line: str) -> bool:
        if len(line) <= ROW_START_COL:
            return False
        return not line[ROW_START_COL].isspace()

    records = []
    current_values = None      # accumulated text per column
    current_last_raw = None    # last raw (padded) segment for each column

    for raw in body_lines:
        if not raw.strip():
            # Completely blank line: ignore.
            continue

        line = pad_to_width(raw, target_len)

        if is_new_record(line):
            # Start a new logical record.
            if current_values is not None:
                # Save the previous record before starting the new one.
                records.append([val.strip() for val in current_values])

            segs = split_line_into_segments(line, spans, target_len)
            current_values = [seg.rstrip() for seg in segs]
            current_last_raw = segs[:]  # copy list
        else:
            # Continuation of the current record.
            if current_values is None:
                # Malformed file: continuation before any record; skip.
                continue

            segs = split_line_into_segments(line, spans, target_len)
            for col_idx, raw_seg in enumerate(segs):
                seg_text = raw_seg.strip()
                if not seg_text:
                    # No new text in this column on this line.
                    continue

                prev_val = current_values[col_idx]

                if not prev_val:
                    # Column was empty so far → just set it.
                    current_values[col_idx] = seg_text
                else:
                    last_raw = current_last_raw[col_idx]
                    width = len(last_raw)
                    trimmed_prev_raw = last_raw.rstrip()

                    # Decide how to join: no space vs single space.
                    if width > 0 and len(trimmed_prev_raw) == width:
                        # Previous piece filled the whole column.
                        # Treat this as mid-word wrapping.
                        current_values[col_idx] = prev_val + seg_text
                    else:
                        # Previous piece ended early; natural break.
                        current_values[col_idx] = prev_val + " " + seg_text

                # Update last raw segment for this column.
                current_last_raw[col_idx] = raw_seg

    # Flush the last record if there is one.
    if current_values is not None:
        records.append([val.strip() for val in current_values])

    return records


def polish_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Lightly polish the DataFrame **values**.

    - Strip leading / trailing whitespace from all string cells.
    - Optionally, collapse internal runs of 2+ spaces to a single space
      for readability. We do this carefully so formulas keep normal spacing.

    You can comment out the internal-space normalization if you want to
    preserve original spacing exactly.
    """

    def clean_cell(x):
        if not isinstance(x, str):
            return x
        # Trim ends
        s = x.strip()
        # Normalize big runs of spaces inside to single spaces (optional, readable).
        s = re.sub(r"\s{2,}", " ", s)
        return s

    return df.applymap(clean_cell)


def style_excel(out_path: str):
    """
    Post-process the written Excel file to make it look nicer:

    - Bold header row
    - Wrap text in all cells
    - Auto-size columns (with a reasonable max width)
    - Freeze header row
    - Add auto-filter on the header row
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment

    wb = load_workbook(out_path)
    ws = wb.active

    # Bold + centered headers, wrap text in header
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Wrap text + top-align for all data cells
    wrap_alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            # Preserve existing horizontal alignment if any; focus on wrap + vertical
            cell.alignment = Alignment(
                horizontal=cell.alignment.horizontal or "left",
                vertical="top",
                wrap_text=True
            )

    # Approximate auto-fit of column widths (limit to max 60 chars)
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            value = cell.value
            if value is None:
                continue
            value_str = str(value)
            if len(value_str) > max_len:
                max_len = len(value_str)
        adjusted_width = min(max_len + 2, 60)
        ws.column_dimensions[col_letter].width = adjusted_width

    # Freeze first row (headers)
    ws.freeze_panes = "A2"

    # Add auto-filter over the used range
    ws.auto_filter.ref = ws.dimensions

    wb.save(out_path)


def main():
    # 1. Ask user to choose the TXT file (panel + optional fallback).
    txt_path = ask_for_file()

    # 2. Read the whole file.
    with open(txt_path, "r", encoding="utf-8", errors="replace") as f:
        lines = [line.rstrip("\n") for line in f]

    if len(lines) < 4:
        print("File is too short (needs at least 3 header lines plus data).")
        sys.exit(1)

    # First three lines are header lines by specification.
    header_lines = lines[:3]
    body_lines = lines[3:]

    # 3. Detect column spans + build column names.
    boundary_line, spans = detect_column_spans(header_lines)
    column_names = build_column_names(header_lines, spans)

    # 4. Parse data lines into logical records.
    records = parse_records(body_lines, boundary_line, spans)

    # 5. Create DataFrame.
    df = pd.DataFrame(records, columns=column_names)

    # 6. Light polish of DataFrame values (optional but nice).
    df = polish_dataframe(df)

    # 7. Save Excel in same folder, same base name + "_clean.xlsx".
    folder = os.path.dirname(os.path.abspath(txt_path))
    base_name = os.path.splitext(os.path.basename(txt_path))[0]
    out_path = os.path.join(folder, base_name + "_clean.xlsx")

    # Use openpyxl engine to write .xlsx.
    df.to_excel(out_path, index=False, engine="openpyxl")

    # 8. Apply Excel styling.
    style_excel(out_path)

    print(f"Converted {len(df)} records.")
    print("Polished Excel file saved as:")
    print(f"  {out_path}")

if __name__ == "__main__":
    main()
