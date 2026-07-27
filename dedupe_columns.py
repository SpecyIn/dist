#!/usr/bin/env python3
"""
Duplicate Column Resolver for Power BI PBIP / TMDL files.
Scans files sequentially for duplicate column definitions, extracts and compares
their lineageTag / sourceLineageTag values, computes and ANSI-color highlights exact line/property
differences between options, queries Git history (git show HEAD) if conflict markers were
already removed (e.g. after accepting both changes), displays whether each option originated
from Current Branch (HEAD) or Incoming Change (branch), tracks remaining duplicate counts per-file
and total across all files, and prompts the user to select which definition to keep (or skip).
"""

import argparse
import difflib
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

# Enable ANSI escape sequences on Windows console
if sys.platform == "win32":
    try:
        os.system("")
    except Exception:
        pass

# ANSI Color Codes for Terminal Highlighting
CLR_RESET = "\033[0m"
CLR_BOLD = "\033[1m"
CLR_YELLOW = "\033[1;33m"
CLR_GREEN = "\033[1;32m"
CLR_RED = "\033[1;31m"
CLR_CYAN = "\033[1;36m"
CLR_BG_YELLOW = "\033[43;30m"


@dataclass
class ColumnBlock:
    col_name: str
    start_line: int  # 0-indexed line number inclusive
    end_line: int    # 0-indexed line number exclusive
    lines: List[str]
    git_side: Optional[str] = None            # "CURRENT" or "INCOMING"
    git_label: Optional[str] = None           # e.g. "Current Branch (HEAD)" or "Incoming Change (feature/branch)"
    conflict_start_line: Optional[int] = None # Line index of <<<<<<< if inside Git conflict
    conflict_sep_line: Optional[int] = None   # Line index of ======= if inside Git conflict
    conflict_end_line: Optional[int] = None   # Line index of >>>>>>> if inside Git conflict

    @property
    def display_text(self) -> str:
        return "".join(self.lines).rstrip()


@dataclass
class FileTask:
    file_path: Path
    has_bom: bool
    lines: List[str]
    duplicate_groups: Dict[str, List[ColumnBlock]]


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

# Keywords that start a new object in TMDL at the same or lower indentation level
OBJECT_KEYWORDS = re.compile(
    r'^\s*(?:column|measure|partition|table|hierarchy|ref|role|expression|perspective|culture)\b',
    re.IGNORECASE
)

# Regex to match lineageTag or sourceLineageTag property lines with optional quotes
TAG_PATTERN = re.compile(
    r'^\s*["\']?(lineageTag|sourceLineageTag)["\']?\s*:\s*(.+)$',
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
        match = TAG_PATTERN.match(line)
        if match:
            key_name = match.group(1).lower()
            val = match.group(2).strip()
            if key_name == "lineagetag":
                lineage_tag = val
            elif key_name == "sourcelineagetag":
                source_lineage_tag = val
    return lineage_tag, source_lineage_tag


def compute_block_differences(col_blocks: List[ColumnBlock]) -> Tuple[List[str], List[Set[int]]]:
    """
    Computes property and line differences between duplicate column blocks.
    Returns:
      diff_summary: list of formatted strings describing property diffs in color.
      differing_lines_per_block: list of sets containing line indices within each block that differ.
    """
    diff_summary: List[str] = []
    differing_lines_per_block: List[Set[int]] = [set() for _ in col_blocks]

    if len(col_blocks) >= 2:
        lines1 = [l.rstrip("\r\n") for l in col_blocks[0].lines]
        lines2 = [l.rstrip("\r\n") for l in col_blocks[1].lines]

        # Compare key-value properties
        props1 = {l.split(":")[0].strip().lower(): l.strip() for l in lines1 if ":" in l}
        props2 = {l.split(":")[0].strip().lower(): l.strip() for l in lines2 if ":" in l}

        all_keys = set(props1.keys()) | set(props2.keys())
        for key in sorted(all_keys):
            v1 = props1.get(key)
            v2 = props2.get(key)
            if v1 != v2:
                if v1 and v2:
                    diff_summary.append(f"{CLR_RED}      - Option [1]: {v1}{CLR_RESET}\n{CLR_GREEN}      + Option [2]: {v2}{CLR_RESET}")
                elif v1:
                    diff_summary.append(f"{CLR_RED}      - Option [1]: {v1}{CLR_RESET}\n{CLR_GREEN}      + Option [2]: (missing property){CLR_RESET}")
                else:
                    diff_summary.append(f"{CLR_RED}      - Option [1]: (missing property){CLR_RESET}\n{CLR_GREEN}      + Option [2]: {v2}{CLR_RESET}")

        # Mark line-by-line differences using difflib
        matcher = difflib.SequenceMatcher(None, [l.strip() for l in lines1], [l.strip() for l in lines2])
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag != 'equal':
                for idx in range(i1, i2):
                    differing_lines_per_block[0].add(idx)
                for idx in range(j1, j2):
                    differing_lines_per_block[1].add(idx)

    return diff_summary, differing_lines_per_block


def format_block_with_diff_highlights(block: ColumnBlock, diff_indices: Set[int]) -> str:
    """
    Formats the lines of a column block, highlighting differing lines with ANSI color, '*' indicator, and tag.
    """
    formatted_lines = []
    for idx, line in enumerate(block.lines):
        clean_line = line.rstrip("\r\n")
        if idx in diff_indices:
            formatted_lines.append(f"{CLR_YELLOW}  * {clean_line}   <-- DIFFERENT{CLR_RESET}")
        else:
            formatted_lines.append(f"    {clean_line}")
    return "\n".join(formatted_lines)


def get_git_head_content(file_path: Path) -> Optional[str]:
    """
    Attempts to retrieve the content of file_path from Git HEAD (or ORIG_HEAD)
    using git CLI commands.
    """
    try:
        repo_root_cmd = ["git", "rev-parse", "--show-toplevel"]
        res = subprocess.run(
            repo_root_cmd,
            cwd=file_path.parent if file_path.is_file() else file_path,
            capture_output=True,
            text=True,
            timeout=5
        )
        if res.returncode != 0:
            return None
        repo_root = Path(res.stdout.strip())

        rel_path = file_path.resolve().relative_to(repo_root.resolve()).as_posix()

        for ref in ("HEAD", "ORIG_HEAD", "MERGE_HEAD"):
            git_show_cmd = ["git", "show", f"{ref}:{rel_path}"]
            res_show = subprocess.run(
                git_show_cmd,
                cwd=repo_root,
                capture_output=True,
                text=True,
                timeout=5
            )
            if res_show.returncode == 0 and res_show.stdout:
                return res_show.stdout
    except Exception:
        pass
    return None


def check_git_origin_for_blocks(
    file_path: Path, col_blocks: List[ColumnBlock]
) -> None:
    """
    If blocks do not already have git_label (i.e. conflict markers were removed),
    queries Git HEAD for the file's previous version to identify which block came from HEAD.
    """
    if any(blk.git_label is not None for blk in col_blocks):
        return

    head_content = get_git_head_content(file_path)
    if not head_content:
        return

    head_lines = head_content.splitlines(keepends=True)
    head_blocks = extract_column_blocks(head_lines)

    col_name_key = col_blocks[0].col_name.lower()
    head_cols = [b for b in head_blocks if b.col_name.lower() == col_name_key]
    if not head_cols:
        return

    head_lt, head_slt = extract_lineage_tags(head_cols[0].lines)

    matched_head = False
    for blk in col_blocks:
        lt, slt = extract_lineage_tags(blk.lines)
        if (head_lt and lt == head_lt) or (head_slt and slt == head_slt):
            blk.git_label = "Current Branch (HEAD)"
            matched_head = True
        elif head_lt is None and head_slt is None and "".join(blk.lines).strip() == "".join(head_cols[0].lines).strip():
            blk.git_label = "Current Branch (HEAD)"
            matched_head = True

    if matched_head:
        for blk in col_blocks:
            if blk.git_label is None:
                blk.git_label = "Incoming Change (feature branch)"


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

        if stripped.startswith("<<<<<<<"):
            in_conflict = True
            current_side = "CURRENT"
            ref = stripped[7:].strip()
            current_label = f"Current Branch ({ref if ref else 'HEAD'})"
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
                    elif curr_indent == header_indent:
                        if OBJECT_KEYWORDS.match(curr_line):
                            break
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


def scan_and_prepare_tasks(
    target_files: List[Path], stats: SummaryStats
) -> Tuple[List[FileTask], int]:
    """
    Pre-scans all target files to filter binary/non-UTF8 files and collect
    files containing duplicate columns along with total duplicate count.
    """
    tasks: List[FileTask] = []
    total_duplicates_all = 0

    for file_path in target_files:
        try:
            with open(file_path, "rb") as f:
                chunk = f.read(8192)
                if b"\x00" in chunk:
                    stats.files_skipped += 1
                    continue
                f.seek(0)
                raw_bytes = f.read()
        except Exception:
            stats.files_skipped += 1
            continue

        has_bom = raw_bytes.startswith(b"\xef\xbb\xbf")
        try:
            if has_bom:
                content = raw_bytes[3:].decode("utf-8")
            else:
                content = raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            stats.files_skipped += 1
            continue

        stats.files_scanned += 1
        lines = content.splitlines(keepends=True)
        blocks = extract_column_blocks(lines)
        duplicate_groups = group_duplicate_columns(blocks)

        if duplicate_groups:
            tasks.append(FileTask(
                file_path=file_path,
                has_bom=has_bom,
                lines=lines,
                duplicate_groups=duplicate_groups
            ))
            total_duplicates_all += len(duplicate_groups)

    return tasks, total_duplicates_all


def process_file_task(
    task: FileTask,
    file_idx: int,
    total_files: int,
    dry_run: bool,
    auto_keep: Optional[str],
    global_remaining: List[int],
    stats: SummaryStats
) -> None:
    """
    Processes duplicate columns for a single file task, displaying highlighted diffs
    and remaining duplicate counts per file and total across all files.
    """
    num_dups_in_file = len(task.duplicate_groups)
    dups_remaining_in_file = num_dups_in_file
    lines_to_delete: Set[int] = set()
    file_modified = False

    print(f"\n{CLR_CYAN}================================================================================{CLR_RESET}")
    print(f"{CLR_BOLD}Processing File [{file_idx}/{total_files}]: {task.file_path}{CLR_RESET}")
    print(f"Found {num_dups_in_file} column(s) with duplicate definitions in this file.")
    print(f"{CLR_CYAN}================================================================================{CLR_RESET}")

    for col_idx, (col_key, col_blocks) in enumerate(task.duplicate_groups.items(), 1):
        check_git_origin_for_blocks(task.file_path, col_blocks)

        col_display_name = col_blocks[0].col_name
        stats.duplicates_found += 1
        num_options = len(col_blocks)

        diff_summary, diff_indices = compute_block_differences(col_blocks)

        if dry_run:
            print(f"\n[DRY-RUN] Column {col_idx}/{num_dups_in_file} ('{col_display_name}') - {num_options} duplicate definitions:")
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
            print(f"\n{CLR_BOLD}Duplicate {col_idx} of {num_dups_in_file} for column: '{col_display_name}'{CLR_RESET}")
            print(f"  --> Remaining in THIS file: {dups_remaining_in_file} | TOTAL remaining across all files: {global_remaining[0]}")
            
            print(f"\n{CLR_CYAN}  • LineageTag Comparison Summary:{CLR_RESET}")
            for idx, blk in enumerate(col_blocks, 1):
                lt, slt = extract_lineage_tags(blk.lines)
                if lt:
                    tag_desc = f"lineageTag: {lt}"
                elif slt:
                    tag_desc = f"sourceLineageTag: {slt}"
                else:
                    tag_desc = "lineageTag: (none found)"

                origin_desc = f"  [{blk.git_label}]" if blk.git_label else f"  [Occurrence {idx}]"
                print(f"      Option [{idx}]: {tag_desc}{origin_desc} (Line {blk.start_line + 1})")

            if diff_summary:
                print(f"\n{CLR_CYAN}  • Key Differences Highlight:{CLR_RESET}")
                for d_line in diff_summary:
                    print(d_line)

            for idx, blk in enumerate(col_blocks, 1):
                origin = f" [{blk.git_label}]" if blk.git_label else f" [Occurrence {idx}]"
                diff_set = diff_indices[idx - 1] if idx - 1 < len(diff_indices) else set()
                formatted_body = format_block_with_diff_highlights(blk, diff_set)

                print(f"\n{CLR_BOLD}--- Option [{idx}]{origin} (Lines {blk.start_line + 1}-{blk.end_line}) ---{CLR_RESET}")
                print(formatted_body)

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

        dups_remaining_in_file -= 1
        global_remaining[0] -= 1

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
        stats.files_modified += 1
        new_lines = [line for idx, line in enumerate(task.lines) if idx not in lines_to_delete]
        new_content = "".join(new_lines)
        encoded_data = new_content.encode("utf-8")
        if task.has_bom:
            encoded_data = b"\xef\xbb\xbf" + encoded_data
        with open(task.file_path, "wb") as f:
            f.write(encoded_data)
        print(f"\n[UPDATED] File updated: {task.file_path}")


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
    target_files: List[Path] = []

    if target_path.is_file():
        target_files.append(target_path)
    elif target_path.is_dir():
        for root, _, files in os.walk(target_path):
            for file in files:
                file_path = Path(root) / file
                if should_process_file(file_path, extensions_set):
                    target_files.append(file_path)
    else:
        print(f"Error: Path '{target_path}' is neither a file nor a directory.", file=sys.stderr)
        sys.exit(1)

    tasks, total_duplicates_all = scan_and_prepare_tasks(target_files, stats)
    global_remaining = [total_duplicates_all]
    total_files = len(tasks)

    for file_idx, task in enumerate(tasks, 1):
        process_file_task(
            task, file_idx, total_files, args.dry_run, auto_keep, global_remaining, stats
        )

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
# python dedupe_columns.py "C:/path/to/file.tmdl"
# python dedupe_columns.py "C:/path/to/folder"
# python dedupe_columns.py "C:/path/to/folder" --dry-run
