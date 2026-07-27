#!/usr/bin/env python3
"""
Duplicate Object Resolver for Power BI PBIP / TMDL files.
Scans files sequentially for duplicate column and/or expression definitions,
extracts and compares their lineageTag / sourceLineageTag values and properties,
computes and ANSI-color highlights exact line/property differences between options,
queries Git history (git show HEAD) if conflict markers were already removed (e.g. after accepting both changes),
displays Incoming Change ALWAYS on top (Option 1) and Current Change (HEAD) on bottom (Option 2),
tracks remaining duplicate counts per-file and total across all files, and prompts the user to select
which definition to keep (or skip).
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
    obj_type: str    # "column", "expression", "measure"
    col_name: str    # object name
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
    duplicate_groups: Dict[Tuple[str, str], List[ColumnBlock]]


@dataclass
class SummaryStats:
    files_scanned: int = 0
    files_modified: int = 0
    duplicates_found: int = 0
    duplicates_resolved: int = 0
    duplicates_skipped: int = 0
    files_skipped: int = 0


# Regex to detect TMDL column, expression, or measure header line
OBJECT_HEADER_PATTERN = re.compile(
    r'^(?P<indent>\s*)(?P<type>column|expression|measure)\s+(?P<header>.+)$',
    re.IGNORECASE
)

# Keywords that start a new object in TMDL at the same or lower indentation level
OBJECT_KEYWORDS = re.compile(
    r'^\s*(?:column|expression|measure|partition|table|hierarchy|ref|role|expression|perspective|culture)\b',
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
    """Extracts object name from TMDL header string, removing expression parts and quotes."""
    name_part = header_str.split("=")[0].strip()
    if len(name_part) >= 2 and name_part[0] in ("'", '"', '`') and name_part[0] == name_part[-1]:
        return name_part[1:-1]
    return name_part


def extract_lineage_tags(lines: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """
    Extracts (lineageTag_value, sourceLineageTag_value) from lines of a block.
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


def sort_blocks_incoming_first(col_blocks: List[ColumnBlock]) -> List[ColumnBlock]:
    """
    Guarantees Incoming Change is ALWAYS Option 1 (on top) and Current Branch (HEAD) is ALWAYS Option 2 (on bottom).
    """
    if len(col_blocks) != 2:
        return col_blocks

    def get_sort_key(block: ColumnBlock) -> int:
        if block.git_side == "INCOMING" or (block.git_label and "Incoming" in block.git_label):
            return 0  # Incoming comes first (Option 1)
        if block.git_side == "CURRENT" or (block.git_label and "Current" in block.git_label):
            return 1  # Current comes second (Option 2)
        return 2

    return sorted(col_blocks, key=get_sort_key)


def compute_block_differences(col_blocks: List[ColumnBlock]) -> Tuple[List[str], List[Set[int]]]:
    """
    Computes exact property and line differences between duplicate object blocks.
    Returns:
      diff_summary: list of formatted strings describing property diffs in color.
      differing_lines_per_block: list of sets containing line indices within each block that differ.
    """
    diff_summary: List[str] = []
    differing_lines_per_block: List[Set[int]] = [set() for _ in col_blocks]

    if len(col_blocks) >= 2:
        lines1_clean = [l.strip() for l in col_blocks[0].lines if l.strip()]
        lines2_clean = [l.strip() for l in col_blocks[1].lines if l.strip()]

        matcher = difflib.SequenceMatcher(None, lines1_clean, lines2_clean)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'replace':
                for a_idx, b_idx in zip(range(i1, i2), range(j1, j2)):
                    diff_summary.append(
                        f"{CLR_RED}      - Option [1]: {lines1_clean[a_idx]}{CLR_RESET}\n"
                        f"{CLR_GREEN}      + Option [2]: {lines2_clean[b_idx]}{CLR_RESET}"
                    )
            elif tag == 'delete':
                for a_idx in range(i1, i2):
                    diff_summary.append(
                        f"{CLR_RED}      - Option [1]: {lines1_clean[a_idx]}{CLR_RESET}\n"
                        f"{CLR_GREEN}      + Option [2]: (missing in Option 2){CLR_RESET}"
                    )
            elif tag == 'insert':
                for b_idx in range(j1, j2):
                    diff_summary.append(
                        f"{CLR_RED}      - Option [1]: (missing in Option 1){CLR_RESET}\n"
                        f"{CLR_GREEN}      + Option [2]: {lines2_clean[b_idx]}{CLR_RESET}"
                    )

        orig_lines1 = [l.rstrip("\r\n") for l in col_blocks[0].lines]
        orig_lines2 = [l.rstrip("\r\n") for l in col_blocks[1].lines]
        orig_matcher = difflib.SequenceMatcher(None, [l.strip() for l in orig_lines1], [l.strip() for l in orig_lines2])
        for tag, i1, i2, j1, j2 in orig_matcher.get_opcodes():
            if tag != 'equal':
                for idx in range(i1, i2):
                    differing_lines_per_block[0].add(idx)
                for idx in range(j1, j2):
                    differing_lines_per_block[1].add(idx)

    return diff_summary, differing_lines_per_block


def format_block_with_diff_highlights(block: ColumnBlock, diff_indices: Set[int]) -> str:
    """
    Formats the lines of an object block, highlighting differing lines with ANSI color, '*' indicator, and tag.
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
    using git CLI commands with case-insensitive Windows path resolution.
    """
    try:
        abs_path = file_path.resolve()
        repo_root_cmd = ["git", "rev-parse", "--show-toplevel"]
        res = subprocess.run(
            repo_root_cmd,
            cwd=abs_path.parent if abs_path.is_file() else abs_path,
            capture_output=True,
            text=True,
            timeout=5
        )
        if res.returncode != 0:
            return None
        repo_root = Path(res.stdout.strip()).resolve()

        abs_path_str = str(abs_path)
        repo_root_str = str(repo_root)

        if len(abs_path_str) >= 2 and abs_path_str[1] == ":":
            abs_path_str = abs_path_str[0].upper() + abs_path_str[1:]
        if len(repo_root_str) >= 2 and repo_root_str[1] == ":":
            repo_root_str = repo_root_str[0].upper() + repo_root_str[1:]

        abs_path_norm = Path(abs_path_str)
        repo_root_norm = Path(repo_root_str)

        try:
            rel_path = abs_path_norm.relative_to(repo_root_norm).as_posix()
        except ValueError:
            rel_path = os.path.relpath(abs_path_norm, repo_root_norm).replace("\\", "/")

        for ref in ("HEAD", "ORIG_HEAD", "MERGE_HEAD"):
            git_show_cmd = ["git", "show", f"{ref}:{rel_path}"]
            res_show = subprocess.run(
                git_show_cmd,
                cwd=repo_root_norm,
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

    target_type = col_blocks[0].obj_type.lower()
    target_name = col_blocks[0].col_name.lower()

    head_matches = [
        b for b in head_blocks
        if b.obj_type.lower() == target_type and b.col_name.lower() == target_name
    ]
    if not head_matches:
        return

    head_lt, head_slt = extract_lineage_tags(head_matches[0].lines)

    matched_head = False
    for blk in col_blocks:
        lt, slt = extract_lineage_tags(blk.lines)
        if (head_lt and lt == head_lt) or (head_slt and slt == head_slt):
            blk.git_side = "CURRENT"
            blk.git_label = "Current Branch (HEAD)"
            matched_head = True
        elif head_lt is None and head_slt is None and "".join(blk.lines).strip() == "".join(head_matches[0].lines).strip():
            blk.git_side = "CURRENT"
            blk.git_label = "Current Branch (HEAD)"
            matched_head = True

    if matched_head:
        for blk in col_blocks:
            if blk.git_label is None:
                blk.git_side = "INCOMING"
                blk.git_label = "Incoming Change (feature branch)"


def extract_column_blocks(lines: List[str]) -> List[ColumnBlock]:
    """
    Parses TMDL file lines and extracts all column, expression, and measure blocks,
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

        match = OBJECT_HEADER_PATTERN.match(line)
        if match:
            indent_str = match.group("indent")
            obj_type = match.group("type").lower()
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
                obj_type=obj_type,
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


def group_duplicate_columns(
    blocks: List[ColumnBlock], target_object_type: str = "all"
) -> Dict[Tuple[str, str], List[ColumnBlock]]:
    """
    Groups object blocks by (obj_type, obj_name.lower()) and filters by target_object_type.
    """
    grouped: Dict[Tuple[str, str], List[ColumnBlock]] = {}
    for block in blocks:
        if target_object_type != "all" and block.obj_type.lower() != target_object_type.lower():
            continue
        key = (block.obj_type.lower(), block.col_name.lower())
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


def get_target_object_type(args_type: Optional[str]) -> str:
    """
    Returns normalized target object type: 'column', 'expression', or 'all'.
    Prompts interactively if not provided as a CLI flag.
    """
    if args_type:
        val = args_type.strip().lower()
        if val in ("1", "column", "columns"):
            return "column"
        elif val in ("2", "expression", "expressions"):
            return "expression"
        elif val in ("3", "all", "both"):
            return "all"

    print("\nSelect target object type to process:")
    print(" 1 - Columns")
    print(" 2 - Expressions")
    print(" 3 - All (Columns & Expressions)")
    while True:
        choice = input("Enter option (1, 2, or 3): ").strip()
        if choice == "1":
            return "column"
        elif choice == "2":
            return "expression"
        elif choice == "3":
            return "all"
        print("Invalid input. Please enter 1, 2, or 3.\n")


def scan_and_prepare_tasks(
    target_files: List[Path], target_object_type: str, stats: SummaryStats
) -> Tuple[List[FileTask], int]:
    """
    Pre-scans all target files to filter binary/non-UTF8 files and collect
    files containing duplicate objects matching target_object_type along with total duplicate count.
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
        duplicate_groups = group_duplicate_columns(blocks, target_object_type)

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
    Processes duplicate objects for a single file task, displaying highlighted diffs,
    ensuring Incoming Change is Option 1 (on top) and Current Change is Option 2 (on bottom),
    and tracking remaining duplicate counts per file and total across all files.
    """
    num_dups_in_file = len(task.duplicate_groups)
    dups_remaining_in_file = num_dups_in_file
    lines_to_delete: Set[int] = set()
    file_modified = False

    print("\n" + CLR_CYAN + "================================================================================" + CLR_RESET)
    print(CLR_BOLD + f"Processing File [{file_idx}/{total_files}]: {task.file_path}" + CLR_RESET)
    print(f"Found {num_dups_in_file} object(s) with duplicate definitions in this file.")
    print(CLR_CYAN + "================================================================================" + CLR_RESET)

    for col_idx, ((obj_type, _), col_blocks) in enumerate(task.duplicate_groups.items(), 1):
        check_git_origin_for_blocks(task.file_path, col_blocks)
        col_blocks = sort_blocks_incoming_first(col_blocks)

        col_display_name = col_blocks[0].col_name
        stats.duplicates_found += 1
        num_options = len(col_blocks)

        diff_summary, diff_indices = compute_block_differences(col_blocks)

        if dry_run:
            print(f"\n[DRY-RUN] {obj_type.capitalize()} {col_idx}/{num_dups_in_file} ('{col_display_name}') - {num_options} duplicate definitions:")
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
            print(f"\n[AUTO-KEEP FIRST] Selected Option 1{label_str} for {obj_type} '{col_display_name}'")
        elif auto_keep == "last":
            selected_idx = num_options - 1
            last_blk = col_blocks[-1]
            label_str = f" ({last_blk.git_label})" if last_blk.git_label else ""
            print(f"\n[AUTO-KEEP LAST] Selected Option {num_options}{label_str} for {obj_type} '{col_display_name}'")
        else:
            print("\n" + CLR_BOLD + f"Duplicate {col_idx} of {num_dups_in_file} for {obj_type}: '{col_display_name}'" + CLR_RESET)
            print(f"  --> Remaining in THIS file: {dups_remaining_in_file} | TOTAL remaining across all files: {global_remaining[0]}")
            
            print("\n" + CLR_CYAN + "  • LineageTag Comparison Summary:" + CLR_RESET)
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
                print("\n" + CLR_CYAN + "  • Key Differences Highlight:" + CLR_RESET)
                for d_line in diff_summary:
                    print(d_line)

            for idx, blk in enumerate(col_blocks, 1):
                origin = f" [{blk.git_label}]" if blk.git_label else f" [Occurrence {idx}]"
                diff_set = diff_indices[idx - 1] if idx - 1 < len(diff_indices) else set()
                formatted_body = format_block_with_diff_highlights(blk, diff_set)

                print("\n" + CLR_BOLD + f"--- Option [{idx}]{origin} (Lines {blk.start_line + 1}-{blk.end_line}) ---" + CLR_RESET)
                print(formatted_body)

            print("-" * 60)
            while True:
                choice = input(
                    f"Select option to KEEP for {obj_type} '{col_display_name}' (1-{num_options}, or 's' to skip): "
                ).strip().lower()

                if choice in ("s", "skip"):
                    print(f"Skipped duplicate {obj_type} '{col_display_name}'.")
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
        description="Resolve duplicate column and/or expression definitions in Power BI PBIP / TMDL files."
    )
    parser.add_argument(
        "path",
        nargs="?",
        help="Path to target file or folder.",
    )
    parser.add_argument(
        "--type",
        choices=["1", "2", "3", "column", "expression", "all"],
        help="Target object type: 1/column, 2/expression, or 3/all.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report duplicate objects without modifying files or prompting.",
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
        help="Automatically keep the first occurrence / Incoming Change without prompting.",
    )
    parser.add_argument(
        "--keep-last",
        action="store_true",
        help="Automatically keep the last occurrence / Current Change without prompting.",
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

    target_object_type = get_target_object_type(args.type)

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

    tasks, total_duplicates_all = scan_and_prepare_tasks(target_files, target_object_type, stats)
    global_remaining = [total_duplicates_all]
    total_files = len(tasks)

    for file_idx, task in enumerate(tasks, 1):
        process_file_task(
            task, file_idx, total_files, args.dry_run, auto_keep, global_remaining, stats
        )

    print("\n--- Summary ---")
    print(f"Files scanned: {stats.files_scanned}")
    print(f"Files modified: {stats.files_modified}")
    print(f"Duplicate object groups found: {stats.duplicates_found}")
    print(f"Duplicate object groups resolved: {stats.duplicates_resolved}")
    print(f"Duplicate object groups skipped: {stats.duplicates_skipped}")
    print(f"Files skipped (binary/non-utf8): {stats.files_skipped}")


if __name__ == "__main__":
    main()


# Example usage:
# python dedupe_columns.py
# python dedupe_columns.py "C:/path/to/folder" --type column
# python dedupe_columns.py "C:/path/to/folder" --type expression
# python dedupe_columns.py "C:/path/to/folder" --type all
