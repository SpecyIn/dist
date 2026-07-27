#!/usr/bin/env python3
"""
Duplicate Column Resolver for Power BI PBIP / TMDL files.
Scans files sequentially for duplicate column definitions, extracts and compares
their lineageTag / sourceLineageTag values, displays whether each option originated
from Current Change (HEAD) or Incoming Change (branch), and prompts the user to select
which definition to keep (or skip).
"""

import argparse
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple


@dataclass
class ColumnBlock:
    col_name: str
    start_line: int  # 0-indexed line number inclusive
    end_line: int    # 0-indexed line number exclusive
    lines: List[str]
    git_side: Optional[str] = None            # "CURRENT" or "INCOMING"
    git_label: Optional[str] = None           # e.g. "Current Change (HEAD)" or "Incoming Change (feature/branch)"
    conflict_start_line: Optional[int] = None # Line index of <<<<<<< if inside Git conflict
    conflict_sep_line: Optional[int] = None   # Line index of ======= if inside Git conflict
    conflict_end_line: Optional[int] = None   # Line index of >>>>>>> if inside Git conflict

    @property
    def display_text(self) -> str:
        return "".join(self.lines).rstrip()


@dataclass
class SummaryStats:
    files_scanned: int = 0
    files_modified: int = 0
    duplicates_found: int = 0
    duplicates_resolved: int = 0
    duplicates_skipped: int = 0
    files_skipped: int = 0


# Regex to detect TMDL column header line (e.g. "column SPECIAL_SENSITIVE_SEC" or "column 'Special Column'")
COLUMN_HEADER_PATTERN = re.compile(r'^(?P<indent>\s*)column\s+(?P<header>.+)$', re.IGNORECASE)

# Regex to match lineageTag or sourceLineageTag property lines
TAG_PATTERN = re.compile(
    r'^\s*(?:"?(lineageTag|sourceLineageTag)"?)\s*:\s*(.+)$',
    re.IGNORECASE
)


def get_indent_level(indent_str: str) -> int:
    """Calculates indentation level converting tabs to 4 spaces."""
    return indent_str.count('\t') * 4 + indent_str.count(' ')


def extract_column_name(header_str: str) -> str:
    """Extracts column name from TMDL header string, removing expression parts and quotes."""
    name_part = header_str.split("=")[0].strip()
    if len(name_part) >= 2 and name_part[0] in ("'", '"', '`') and name_part[0] == name_part[-1]:
        return name_part[1:-1]
    return name_part


def extract_lineage_tags(lines: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """
    Extracts (lineageTag_value, sourceLineageTag_value) from lines of a column block.
    """
    lineage_tag = None
    source_lineage_tag = None
    for line in lines:
        match = TAG_PATTERN.match(line.strip())
        if match:
            key_name = match.group(1).lower()
            val = match.group(2).strip()
            if key_name == "lineagetag":
                lineage_tag = val
            elif key_name == "sourcelineagetag":
                source_lineage_tag = val
    return lineage_tag, source_lineage_tag


def extract_column_blocks(lines: List[str]) -> List[ColumnBlock]:
    """
    Parses TMDL file lines and extracts all column definition blocks,
    tracking whether each block is inside a Git merge conflict (HEAD vs incoming branch).
    """
    blocks: List[ColumnBlock] = []
    i = 0
    n = len(lines)

    in_conflict = False
    current_side: Optional[str] = None
    current_label: Optional[str] = None
    conflict_start_idx: Optional[int] = None
    conflict_sep_idx: Optional[int] = None
    conflict_end_idx: Optional[int] = None
    active_conflict_blocks: List[ColumnBlock] = []

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # Check for Git conflict markers
        if stripped.startswith("<<<<<<<"):
            in_conflict = True
            current_side = "CURRENT"
            ref = stripped[7:].strip()
            current_label = f"Current Change ({ref if ref else 'HEAD'})"
            conflict_start_idx = i
            conflict_sep_idx = None
            conflict_end_idx = None
            active_conflict_blocks = []
            i += 1
            continue
        elif stripped.startswith("=======") and in_conflict:
            current_side = "INCOMING"
            conflict_sep_idx = i
            current_label = "Incoming Change"
            i += 1
            continue
        elif stripped.startswith(">>>>>>>") and in_conflict:
            ref = stripped[7:].strip()
            inc_label = f"Incoming Change ({ref if ref else 'incoming'})"
            conflict_end_idx = i
            for blk in active_conflict_blocks:
                if blk.git_side == "INCOMING":
                    blk.git_label = inc_label
                blk.conflict_start_line = conflict_start_idx
                blk.conflict_sep_line = conflict_sep_idx
                blk.conflict_end_line = conflict_end_idx

            in_conflict = False
            current_side = None
            current_label = None
            active_conflict_blocks = []
            i += 1
            continue

        match = COLUMN_HEADER_PATTERN.match(line)
        if match:
            indent_str = match.group("indent")
            header_str = match.group("header")

            col_name = extract_column_name(header_str)
            header_indent = get_indent_level(indent_str)
            start_idx = i
            j = i + 1

            while j < n:
                curr_line = lines[j]
                curr_stripped = curr_line.strip()

                if curr_stripped.startswith("=======") or curr_stripped.startswith(">>>>>>>") or curr_stripped.startswith("<<<<<<<"):
                    break

                if not curr_stripped:
                    k = j + 1
                    next_indent = None
                    while k < n:
                        k_stripped = lines[k].strip()
                        if k_stripped.startswith("=======") or k_stripped.startswith(">>>>>>>") or k_stripped.startswith("<<<<<<<"):
                            break
                        if k_stripped:
                            m = re.match(r"^\s*", lines[k])
                            next_indent = get_indent_level(m.group(0)) if m else 0
                            break
                        k += 1

                    if next_indent is not None and next_indent <= header_indent:
                        break
                    j += 1
                else:
                    m = re.match(r"^\s*", curr_line)
                    curr_indent = get_indent_level(m.group(0)) if m else 0

                    if curr_indent > header_indent:
                        j += 1
                    else:
                        break

            block_lines = lines[start_idx:j]
            block = ColumnBlock(
                col_name=col_name,
                start_line=start_idx,
                end_line=j,
                lines=block_lines,
                git_side=current_side,
                git_label=current_label,
            )

            if in_conflict:
                active_conflict_blocks.append(block)

            blocks.append(block)
            i = j
        else:
            i += 1

    return blocks


def group_duplicate_columns(blocks: List[ColumnBlock]) -> Dict[str, List[ColumnBlock]]:
    """Groups column blocks by case-insensitive column name."""
    grouped: Dict[str, List[ColumnBlock]] = {}
    for block in blocks:
        key = block.col_name.lower()
        grouped.setdefault(key, []).append(block)

    return {k: v for k, v in grouped.items() if len(v) > 1}


def resolve_file_duplicates(
    file_path: Path,
    lines: List[str],
    dry_run: bool,
    auto_keep: Optional[str],
    stats: SummaryStats
) -> Tuple[List[str], bool]:
    """
    Checks a single file for duplicate column blocks and resolves them interactively or automatically.
    Displays lineageTag values and Current/Incoming origin for clear comparison during selection.
    """
    blocks = extract_column_blocks(lines)
    duplicate_groups = group_duplicate_columns(blocks)

    if not duplicate_groups:
        return lines, False

    lines_to_delete: Set[int] = set()
    file_modified = False

    print(f"\n================================================================================")
    print(f"Processing File: {file_path}")
    print(f"Found {len(duplicate_groups)} column(s) with duplicate definitions.")
    print(f"================================================================================")

    for col_key, col_blocks in duplicate_groups.items():
        col_display_name = col_blocks[0].col_name
        stats.duplicates_found += 1
        num_options = len(col_blocks)

        if dry_run:
            print(f"\n[DRY-RUN] Column '{col_display_name}' has {num_options} duplicate definitions:")
            for idx, blk in enumerate(col_blocks, 1):
                lt, slt = extract_lineage_tags(blk.lines)
                tag_info = f" | lineageTag: {lt}" if lt else f" | sourceLineageTag: {slt}" if slt else ""
                label_info = f" [{blk.git_label}]" if blk.git_label else f" [Occurrence {idx}]"
                print(f"  Option [{idx}]{label_info}{tag_info}: lines {blk.start_line + 1}-{blk.end_line}")
            continue

        selected_idx: Optional[int] = None

        if auto_keep == "first":
            selected_idx = 0
            label_str = f" ({col_blocks[0].git_label})" if col_blocks[0].git_label else ""
            print(f"\n[AUTO-KEEP FIRST] Selected Option 1{label_str} for column '{col_display_name}'")
        elif auto_keep == "last":
            selected_idx = num_options - 1
            last_blk = col_blocks[-1]
            label_str = f" ({last_blk.git_label})" if last_blk.git_label else ""
            print(f"\n[AUTO-KEEP LAST] Selected Option {num_options}{label_str} for column '{col_display_name}'")
        else:
            print(f"\nDuplicate definitions for column: '{col_display_name}'")
            print("  • LineageTag Comparison Summary:")

            for idx, blk in enumerate(col_blocks, 1):
                lt, slt = extract_lineage_tags(blk.lines)
                if lt:
                    tag_desc = f"lineageTag: {lt}"
                elif slt:
                    tag_desc = f"sourceLineageTag: {slt}"
                else:
                    tag_desc = "lineageTag: (none)"

                origin_desc = f"  [{blk.git_label}]" if blk.git_label else ""
                print(f"      Option [{idx}]: {tag_desc}{origin_desc} (Line {blk.start_line + 1})")

            for idx, blk in enumerate(col_blocks, 1):
                origin = f" [{blk.git_label}]" if blk.git_label else f" [Occurrence {idx}]"
                print(f"\n--- Option [{idx}]{origin} (Lines {blk.start_line + 1}-{blk.end_line}) ---")
                print(blk.display_text)

            print("-" * 60)
            while True:
                choice = input(
                    f"Select option to KEEP for '{col_display_name}' (1-{num_options}, or 's' to skip): "
                ).strip().lower()

                if choice in ("s", "skip"):
                    print(f"Skipped duplicate column '{col_display_name}'.")
                    stats.duplicates_skipped += 1
                    break
                elif choice.isdigit():
                    val = int(choice)
                    if 1 <= val <= num_options:
                        selected_idx = val - 1
                        break
                print(f"Invalid input. Please enter a number between 1 and {num_options}, or 's'.")

        if selected_idx is not None:
            stats.duplicates_resolved += 1
            file_modified = True
            for idx, blk in enumerate(col_blocks):
                if idx != selected_idx:
                    for line_no in range(blk.start_line, blk.end_line):
                        lines_to_delete.add(line_no)

            for blk in col_blocks:
                if blk.conflict_start_line is not None:
                    lines_to_delete.add(blk.conflict_start_line)
                if blk.conflict_sep_line is not None:
                    lines_to_delete.add(blk.conflict_sep_line)
                if blk.conflict_end_line is not None:
                    lines_to_delete.add(blk.conflict_end_line)

    if file_modified and not dry_run:
        new_lines = [line for idx, line in enumerate(lines) if idx not in lines_to_delete]
        return new_lines, True

    return lines, False


def process_file(
    file_path: Path,
    dry_run: bool,
    auto_keep: Optional[str],
    stats: SummaryStats
) -> None:
    """Reads, processes, and writes a file if duplicate columns were resolved."""
    try:
        with open(file_path, "rb") as f:
            chunk = f.read(8192)
            if b"\x00" in chunk:
                stats.files_skipped += 1
                return
            f.seek(0)
            raw_bytes = f.read()
    except Exception:
        stats.files_skipped += 1
        return

    has_bom = raw_bytes.startswith(b"\xef\xbb\xbf")
    try:
        if has_bom:
            content = raw_bytes[3:].decode("utf-8")
        else:
            content = raw_bytes.decode("utf-8")
    except UnicodeDecodeError:
        stats.files_skipped += 1
        return

    stats.files_scanned += 1
    lines = content.splitlines(keepends=True)

    new_lines, was_modified = resolve_file_duplicates(
        file_path, lines, dry_run, auto_keep, stats
    )

    if was_modified:
        stats.files_modified += 1
        if not dry_run:
            new_content = "".join(new_lines)
            encoded_data = new_content.encode("utf-8")
            if has_bom:
                encoded_data = b"\xef\xbb\xbf" + encoded_data
            with open(file_path, "wb") as f:
                f.write(encoded_data)
            print(f"\n[UPDATED] File updated: {file_path}")
        else:
            print(f"\n[DRY-RUN] Would update file: {file_path}")


def parse_extensions(ext_str: Optional[str]) -> Optional[Set[str]]:
    """Parses comma-separated extension string into a normalized set."""
    if not ext_str:
        return None
    exts = {e.strip().lstrip(".").lower() for e in ext_str.split(",") if e.strip()}
    return exts if exts else None


def should_process_file(file_path: Path, extensions: Optional[Set[str]]) -> bool:
    """Checks if file extension matches the filter."""
    if extensions is None:
        return True
    ext = file_path.suffix.lstrip(".").lower()
    return ext in extensions


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Resolve duplicate column definitions in Power BI PBIP / TMDL files."
    )
    parser.add_argument(
        "path",
        nargs="?",
        help="Path to target file or folder.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report duplicate columns without modifying files or prompting.",
    )
    parser.add_argument(
        "--extensions",
        type=str,
        default="tmdl,json,pbir,pbip",
        help="Comma-separated file extensions to process (default: tmdl,json,pbir,pbip). Use 'all' for all text files.",
    )
    parser.add_argument(
        "--keep-first",
        action="store_true",
        help="Automatically keep the first occurrence / Current Change without prompting.",
    )
    parser.add_argument(
        "--keep-last",
        action="store_true",
        help="Automatically keep the last occurrence / Incoming Change without prompting.",
    )

    args = parser.parse_args()

    path_str = args.path
    if not path_str:
        path_str = input("Enter file or folder path: ").strip()

    path_str = path_str.strip('"\'')
    if not path_str:
        print("Error: No path provided.", file=sys.stderr)
        sys.exit(1)

    target_path = Path(path_str)
    if not target_path.exists():
        print(f"Error: Path '{path_str}' does not exist.", file=sys.stderr)
        sys.exit(1)

    ext_arg = None if args.extensions and args.extensions.lower() == "all" else args.extensions
    extensions_set = parse_extensions(ext_arg)

    auto_keep = None
    if args.keep_first:
        auto_keep = "first"
    elif args.keep_last:
        auto_keep = "last"

    stats = SummaryStats()

    if target_path.is_file():
        if should_process_file(target_path, extensions_set):
            process_file(target_path, args.dry_run, auto_keep, stats)
        else:
            print(f"Skipped (extension mismatch): {target_path}")
    elif target_path.is_dir():
        for root, _, files in os.walk(target_path):
            for file in files:
                file_path = Path(root) / file
                if should_process_file(file_path, extensions_set):
                    process_file(file_path, args.dry_run, auto_keep, stats)
    else:
        print(f"Error: Path '{target_path}' is neither a file nor a directory.", file=sys.stderr)
        sys.exit(1)

    print("\n--- Summary ---")
    print(f"Files scanned: {stats.files_scanned}")
    print(f"Files modified: {stats.files_modified}")
    print(f"Duplicate column groups found: {stats.duplicates_found}")
    print(f"Duplicate column groups resolved: {stats.duplicates_resolved}")
    print(f"Duplicate column groups skipped: {stats.duplicates_skipped}")
    print(f"Files skipped (binary/non-utf8): {stats.files_skipped}")


if __name__ == "__main__":
    main()


# Example usage:
# python dedupe_columns.py
# python dedupe_columns.py "C:/path/to/folder"
# python dedupe_columns.py "C:/path/to/folder" --dry-run
# python dedupe_columns.py "C:/path/to/folder" --extensions tmdl,json
