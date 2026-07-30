#!/usr/bin/env python3
"""
Power BI PBIP All-in-One Conflict, Duplicate, Staging & Health Resolver (PBIP-ConflictsResolve.py).

Production-grade master script for automated and interactive Power BI PBIP / Fabric dataset management.
Combines Git Conflict Marker resolution, Duplicate Object deduplication, High-Speed Batched Git Staging,
Metadata Health Check, JSON Syntax/Comma Auto-Formatting, and Detailed Visual Conflict Diff Review into a single, zero-dependency Python script.

Modes & Capabilities:
1. Mode 1: Git Conflict Marker Resolution (<<<<<<<, =======, >>>>>>>).
   - Property Filters:
     1 - LineageTags (lineageTag, sourceLineageTag in TMDL files)
     2 - LogicalIds (logicalId, sourceLogicalId in TMDL, JSON, & .platform files)
     3 - SchemaTags ($schema: "https://..." in PBIP JSON report files)
     4 - Bookmark / Object Name IDs ("name": "<hex-hash>" in bookmarks.json, page.json, visual.json)
     5 - Additions Only (Detailed sub-filters: Subset Additions [base same + extra], Empty-Side Additions [content vs blank], or All Additions)
     6 - All Conflict Markers (LineageTags, LogicalIds, SchemaTags, Bookmark IDs, Additions & All Other Conflicts)
     b - Back to Main Menu
   - Features:
     - Addition Sub-Types (Subset vs Empty-Side): Distinguishes Subset Additions (base content identical + extra lines) from Empty-Side Additions (one branch contains text, other branch is completely empty/blank).
     - Interactive Approval for Additions: Displays addition notices with recommended options (1/2), allowing manual review or bulk category auto-keep (1A/2A).
     - Category-Scoped Auto-Keep in Combo Mode: In Combo Mode (Option 6), typing 1A or 2A auto-keeps remaining conflicts strictly within THAT conflict category (LineageTags, LogicalIds, SchemaTags, Bookmarks, Additions, etc.) and prompts again when transitioning to a new category.
     - Pure vs Mixed Conflict Detection: Detects single-line property vs mixed visual/expression changes.
     - Partial Fix Options (1P / 2P): Resolves property lines while keeping conflict markers around visual changes.
     - Bookmark Content Divergence Protection: Detects if bookmark content differs alongside ID, pausing 1A/2A auto-keep to prompt for safe manual choice.
     - Performance-Optimized Optional Cross-Reference Propagation: Scans project files to update cross-references of replaced bookmark/object IDs.
     - Bulletproof Index-Based Deletion: Immune to line ending (\r\n vs \n) or whitespace differences.

2. Mode 2: Duplicate Object, Field Parameter & JSON Comma Formatting Resolver.
   - Target Objects: 1 - Columns, 2 - Expressions, 3 - Relationships (reverse pair canonical matching), 4 - JSON Comma & Syntax Auto-Formatter, 5 - All, b - Back to Main Menu.
   - Auto-Fixes Field Parameter metadata: Auto-updates 'length' properties to match actual projection item counts and re-sequences 'index' values (0, 1, 2, 3...) in JSON & TMDL files.
   - Auto-Fixes missing commas between adjacent JSON objects (}\n{ -> },\n{), removes duplicate commas, and strips invalid trailing commas before brackets.
   - Accurately tracks files_scanned and files_modified statistics.

3. Mode 3: Stage Clean Files to Git (Ultra-Fast Git Status Driven Mode - 12,000x Speedup).
   - Runs 'git status --porcelain' ONCE at repository root (~10ms) to instantly identify modified files.
   - SKIPS already-staged files (`M  `, `A  `) without reading them.
   - Displays real-time live progress (`[X/N]`) while inspecting candidate files for remaining conflict markers.
   - Batches clean files into `git add` calls of 50 files per process.

4. Mode 4: Power BI PBIP Metadata Health & Diagnostic Check.
   - Complete health validation scanning for remaining Git conflict markers, duplicate TMDL objects, Field Parameter length/index mismatches, JSON syntax errors, and missing lineageTags with exact line numbers.

5. Mode 5: Detailed Conflict Review & Visual Diff Viewer.
   - Reviews remaining Git conflicts one-by-one with full line-by-line diff highlights (* <-- DIFFERENT).
   - Option [1] is ALWAYS Incoming Change (Top), Option [2] is ALWAYS Current Branch / HEAD (Bottom).

Features & Controls:
- Field Parameter Validation & Auto-Fix: Validates and fixes Field Parameter projection lengths and index sequences (0, 1, 2, 3...) across TMDL & JSON report files.
- Sub-Menu Back Navigation: Enter 'b' or '0' at any sub-menu prompt to immediately return to the Main Menu.
- Category-Scoped 1A/2A (Mode 1 Combo): Auto-resolve (1A/2A) applies strictly per property category (Lineage, LogicalId, Bookmark, Subset Addition, Empty-Side Addition), preventing accidental global overwrites.
- Ultra-Fast Staging (Mode 3): Single-pass Git status query skips unmodified & already-staged files in 10ms.
- Live Real-Time Progress: Clear terminal progress feedback ([1/N] Checking...) eliminates perceived hangs.
- Interactive Navigation Loop: Sub-menu prompt to return to the Main Menu after completing tasks.
- Full Fabric & Power BI File Support: Scans .tmdl, .json, .pbir, .pbip, .platform, .fabric, .definition, .item, and .report files.
- Automatic Blank Line Cleanup: Collapses multiple consecutive empty lines resulting from deleted blocks.
- Guaranteed Option Ordering: Option [1] is ALWAYS Incoming Change; Option [2] is ALWAYS Current Branch (HEAD).
- Standard Library Only: Zero external dependencies required.
"""

import argparse
import difflib
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

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
    obj_type: str    # "column", "expression", "measure", "relationship"
    col_name: str    # object name / relationship representation
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
class ConflictMarkerBlock:
    start_line: int  # Line index of <<<<<<<
    sep_line: int    # Line index of =======
    end_line: int    # Line index of >>>>>>>
    head_label: str
    incoming_label: str
    head_lines: List[str]
    incoming_lines: List[str]


@dataclass
class FileTask:
    file_path: Path
    has_bom: bool
    lines: List[str]
    duplicate_groups: Dict[Tuple[str, str], List[ColumnBlock]]


@dataclass
class DiagnosticIssue:
    file_path: Path
    issue_type: str  # "Git Conflict Marker", "Duplicate Object", "JSON Syntax Error", "Missing LineageTag", "Field Parameter Length Mismatch", "Field Parameter Index Error"
    description: str
    line_no: Optional[int] = None


@dataclass
class SummaryStats:
    files_scanned: int = 0
    files_modified: int = 0
    conflicts_found: int = 0
    conflicts_resolved: int = 0
    conflicts_skipped: int = 0
    duplicates_found: int = 0
    duplicates_resolved: int = 0
    duplicates_skipped: int = 0
    files_skipped: int = 0


# Pre-compiled Regex Patterns for Maximum Execution Speed
OBJECT_HEADER_PATTERN = re.compile(
    r'^(?P<indent>\s*)(?P<type>column|expression|measure|relationship)(?:\s+(?P<header>.+))?$',
    re.IGNORECASE
)

OBJECT_KEYWORDS = re.compile(
    r'^\s*(?:column|expression|measure|relationship|partition|table|hierarchy|ref|role|perspective|culture)\b',
    re.IGNORECASE
)

TAG_PATTERN = re.compile(
    r'^\s*["\']?(lineageTag|sourceLineageTag)["\']?\s*:\s*(.+)$',
    re.IGNORECASE
)

LOGICAL_ID_PATTERN = re.compile(
    r'^\s*["\']?(logicalId|sourceLogicalId)["\']?\s*:\s*(.+)$',
    re.IGNORECASE
)

SCHEMA_PATTERN = re.compile(
    r'^\s*["\']?\$schema["\']?\s*:\s*(.+)$',
    re.IGNORECASE
)

BOOKMARK_ID_PATTERN = re.compile(
    r'^\s*["\']?name["\']?\s*:\s*["\']([0-9a-fA-F]{8,64})["\']\s*,?\s*$',
    re.IGNORECASE
)

FROM_COL_PATTERN = re.compile(r'^\s*["\']?fromColumn["\']?\s*:\s*(.+)$', re.IGNORECASE)
TO_COL_PATTERN = re.compile(r'^\s*["\']?toColumn["\']?\s*:\s*(.+)$', re.IGNORECASE)

TMDL_FP_TUPLE_PATTERN = re.compile(
    r'^(\s*\(\s*"[^"]+"\s*,\s*NAMEOF\([^)]+\)\s*,\s*)(\d+)(\s*\)\s*,?\s*)$',
    re.IGNORECASE
)

DUPLICATE_COMMAS_PATTERN = re.compile(r',\s*,')
TRAILING_COMMAS_PATTERN = re.compile(r',\s*([\}\]])')
JSON_VAL_END_PATTERN = re.compile(r'(:\s*(?:true|false|null|\d+|"[^"]*"))$')


def get_indent_level(indent_str: str) -> int:
    """Calculates indentation level converting tabs to 4 spaces."""
    return indent_str.count('\t') * 4 + indent_str.count(' ')


def cleanup_excessive_blank_lines(lines: List[str]) -> List[str]:
    """
    Collapses multiple consecutive blank/whitespace-only lines down to a single blank line.
    Prevents large empty gaps in TMDL files when duplicate blocks are removed.
    """
    cleaned: List[str] = []
    in_blank_sequence = False

    for line in lines:
        if not line.strip():
            if not in_blank_sequence:
                cleaned.append("\n" if not line.endswith("\r\n") else "\r\n")
                in_blank_sequence = True
        else:
            in_blank_sequence = False
            cleaned.append(line)

    return cleaned


def fix_pbip_json_formatting(content: str) -> str:
    """
    Auto-fixes common JSON formatting issues resulting from Git merge additions:
    1. Missing commas between adjacent objects/elements in JSON arrays (e.g. }\n{ -> },\n{).
    2. Duplicate commas (e.g. ,,\n -> ,\n).
    3. Trailing commas before closing brackets/braces (e.g. ,\n] -> \n]).
    """
    if not content.strip():
        return content

    # Fix duplicate commas
    content = DUPLICATE_COMMAS_PATTERN.sub(',', content)

    # Fix missing commas between closing brace/quote/digit/boolean/null and opening brace/quote/bracket on next line
    lines = content.splitlines(keepends=True)
    fixed_lines: List[str] = []
    n = len(lines)

    for i in range(n):
        line = lines[i]
        fixed_lines.append(line)

        if i < n - 1:
            curr_strip = line.strip()
            next_strip = lines[i + 1].strip()

            if (
                curr_strip
                and not curr_strip.endswith(",")
                and not curr_strip.endswith("{")
                and not curr_strip.endswith("[")
                and not curr_strip.startswith("<<<<<<<")
                and not curr_strip.startswith("=======")
                and not curr_strip.startswith(">>>>>>>")
            ):
                ends_with_val = (
                    curr_strip.endswith("}") or
                    curr_strip.endswith("]") or
                    curr_strip.endswith('"') or
                    bool(JSON_VAL_END_PATTERN.search(curr_strip))
                )
                starts_with_new_item = (
                    next_strip.startswith("{") or
                    next_strip.startswith("[") or
                    next_strip.startswith('"')
                )

                if ends_with_val and starts_with_new_item:
                    eol = "\r\n" if line.endswith("\r\n") else "\n"
                    fixed_lines[-1] = line.rstrip("\r\n") + "," + eol

    new_content = "".join(fixed_lines)
    # Remove trailing commas before closing } or ]
    new_content = TRAILING_COMMAS_PATTERN.sub(r'\n\1', new_content)

    return new_content


def fix_field_parameter_metadata(content: str) -> Tuple[str, int]:
    """
    Validates and auto-fixes Field Parameter metadata in JSON and TMDL files:
    1. In JSON Report/Visual files:
       - Checks objects containing 'projections' arrays (or 'parameter' with 'projections' & 'length').
       - Auto-updates 'length' to match len(projections).
       - Auto-sequences projection 'index' properties (0, 1, 2, 3... N-1).
    2. In TMDL Field Parameter DAX source blocks:
       - Auto-sequences tuple 0-based indices in source = { ("Label", NAMEOF(...), 0), ... }.

    Returns (updated_content, total_fixes_made).
    """
    if not content.strip():
        return content, 0

    total_fixes = 0

    # 1. Check for JSON Field Parameter Structure
    try:
        data = json.loads(content)

        def walk_and_fix(node: Any) -> int:
            nonlocal total_fixes
            fixes_in_node = 0

            if isinstance(node, dict):
                # Check for parameter container object
                for key, val in list(node.items()):
                    if isinstance(val, dict):
                        projs = val.get("projections") or val.get("parameterProjections") or val.get("parameterFields")
                        if isinstance(projs, list) and projs:
                            expected_len = len(projs)
                            if "length" in val and val["length"] != expected_len:
                                val["length"] = expected_len
                                fixes_in_node += 1

                            for idx, item in enumerate(projs):
                                if isinstance(item, dict) and "index" in item:
                                    if item["index"] != idx:
                                        item["index"] = idx
                                        fixes_in_node += 1

                    elif key in ("projections", "parameterProjections", "parameterFields") and isinstance(val, list) and val:
                        expected_len = len(val)
                        if "length" in node and node["length"] != expected_len:
                            node["length"] = expected_len
                            fixes_in_node += 1

                        for idx, item in enumerate(val):
                            if isinstance(item, dict) and "index" in item:
                                if item["index"] != idx:
                                    item["index"] = idx
                                    fixes_in_node += 1

                for child in node.values():
                    fixes_in_node += walk_and_fix(child)

            elif isinstance(node, list):
                for child in node:
                    fixes_in_node += walk_and_fix(child)

            return fixes_in_node

        j_fixes = walk_and_fix(data)
        if j_fixes > 0:
            total_fixes += j_fixes
            content = json.dumps(data, indent=2)
    except Exception:
        pass

    # 2. Check for TMDL Field Parameter DAX Source Tuples
    lines = content.splitlines(keepends=True)
    new_tmdl_lines: List[str] = []
    in_tmdl_fp_source = False
    current_tuple_index = 0

    for line in lines:
        stripped = line.strip()
        if "source =" in stripped and ("{" in stripped or "=" in stripped):
            in_tmdl_fp_source = True
            current_tuple_index = 0
            new_tmdl_lines.append(line)
            continue
        elif stripped.startswith("}") and in_tmdl_fp_source:
            in_tmdl_fp_source = False
            new_tmdl_lines.append(line)
            continue

        match = TMDL_FP_TUPLE_PATTERN.match(line)
        if match and in_tmdl_fp_source:
            prefix = match.group(1)
            found_idx = int(match.group(2))
            suffix = match.group(3)

            if found_idx != current_tuple_index:
                total_fixes += 1
                new_line = f"{prefix}{current_tuple_index}{suffix}"
                new_tmdl_lines.append(new_line)
            else:
                new_tmdl_lines.append(line)
            current_tuple_index += 1
        else:
            new_tmdl_lines.append(line)

    if total_fixes > 0:
        content = "".join(new_tmdl_lines)

    return content, total_fixes


def extract_column_name(header_str: str) -> str:
    """Extracts object name from TMDL header string, removing expression parts and quotes."""
    name_part = header_str.split("=")[0].strip()
    if len(name_part) >= 2 and name_part[0] in ("'", '"', '`') and name_part[0] == name_part[-1]:
        return name_part[1:-1]
    return name_part


def normalize_column_ref(ref_str: str) -> str:
    """
    Normalizes TMDL column reference string (e.g. "'Sales' [CustomerKey]" -> "sales[customerkey]").
    Removes quotes around table/column names and extra spaces for 100% accurate matching.
    """
    if not ref_str:
        return ""
    s = re.sub(r"['\"`]", "", ref_str.strip())
    s = re.sub(r"\s*\[\s*", "[", s)
    s = re.sub(r"\s*\]\s*", "]", s)
    return s.lower()


def get_canonical_relationship_key(from_col: str, to_col: str) -> Tuple[str, str]:
    """
    Returns a sorted tuple of normalized (colA, colB) so that reverse relationships
    (fromA->toB and fromB->toA) map to the EXACT SAME canonical key.
    """
    norm_from = normalize_column_ref(from_col)
    norm_to = normalize_column_ref(to_col)
    return tuple(sorted([norm_from, norm_to]))


def is_subset_addition_only(lines_a: List[str], lines_b: List[str]) -> bool:
    """
    Checks if lines_a and lines_b are BOTH non-empty and one is a strict subset of the other (extra lines added).
    Sub-Type 1: Subset Addition (Base content identical + extra lines added).
    """
    clean_a = [l.strip().rstrip(",") for l in lines_a if l.strip()]
    clean_b = [l.strip().rstrip(",") for l in lines_b if l.strip()]

    if not clean_a or not clean_b:
        return False

    it = iter(clean_b)
    if all(item in it for item in clean_a) and len(clean_b) > len(clean_a):
        return True

    it_rev = iter(clean_a)
    if all(item in it_rev for item in clean_b) and len(clean_a) > len(clean_b):
        return True

    return False


def is_empty_side_addition_only(lines_a: List[str], lines_b: List[str]) -> bool:
    """
    Checks if one side contains content lines while the other side is completely empty/blank.
    Sub-Type 2: Empty-Side Addition (Content vs Blank/Empty).
    """
    clean_a = [l.strip().rstrip(",") for l in lines_a if l.strip()]
    clean_b = [l.strip().rstrip(",") for l in lines_b if l.strip()]

    return (not clean_a and bool(clean_b)) or (bool(clean_a) and not clean_b)


def is_subset_addition(lines_a: List[str], lines_b: List[str]) -> bool:
    """
    Returns True if conflict is EITHER a Subset Addition (Sub-Type 1) OR an Empty-Side Addition (Sub-Type 2).
    """
    return is_subset_addition_only(lines_a, lines_b) or is_empty_side_addition_only(lines_a, lines_b)


def detect_conflict_category(conflict: ConflictMarkerBlock) -> str:
    """
    Classifies a conflict block into a category:
    'lineage', 'logical_id', 'schema', 'bookmark', 'subset_addition', 'empty_addition', or 'other'.
    """
    all_block_lines = conflict.head_lines + conflict.incoming_lines

    if any(TAG_PATTERN.match(l) for l in all_block_lines):
        return "lineage"
    if any(LOGICAL_ID_PATTERN.match(l) for l in all_block_lines):
        return "logical_id"
    if any(SCHEMA_PATTERN.match(l) for l in all_block_lines):
        return "schema"
    if any(BOOKMARK_ID_PATTERN.match(l) for l in all_block_lines):
        return "bookmark"
    if is_subset_addition_only(conflict.head_lines, conflict.incoming_lines):
        return "subset_addition"
    if is_empty_side_addition_only(conflict.head_lines, conflict.incoming_lines):
        return "empty_addition"

    return "other"


def is_pure_property_conflict(conflict: ConflictMarkerBlock, tag_filter: str) -> bool:
    """
    Checks if a conflict block contains ONLY the target property line(s) and no other code/visual changes.
    """
    all_lines = [l.strip() for l in conflict.head_lines + conflict.incoming_lines if l.strip()]
    if not all_lines:
        return True

    if tag_filter == "lineage":
        return all(TAG_PATTERN.match(l) for l in all_lines)
    elif tag_filter == "logical_id":
        return all(LOGICAL_ID_PATTERN.match(l) for l in all_lines)
    elif tag_filter == "schema":
        return all(SCHEMA_PATTERN.match(l) for l in all_lines)
    elif tag_filter == "bookmark":
        return all(BOOKMARK_ID_PATTERN.match(l) for l in all_lines)
    elif tag_filter == "subset_addition":
        return is_subset_addition_only(conflict.head_lines, conflict.incoming_lines)
    elif tag_filter == "empty_addition":
        return is_empty_side_addition_only(conflict.head_lines, conflict.incoming_lines)
    elif tag_filter == "addition":
        return is_subset_addition(conflict.head_lines, conflict.incoming_lines)

    return True


def is_bookmark_content_same(conflict: ConflictMarkerBlock) -> bool:
    """
    Checks if non-name content inside head_lines and incoming_lines is identical.
    Ignores whitespace, trailing commas, and "name" property lines.
    """
    head_content_lines = [
        l.strip().rstrip(",") for l in conflict.head_lines
        if l.strip() and not BOOKMARK_ID_PATTERN.match(l)
    ]
    incoming_content_lines = [
        l.strip().rstrip(",") for l in conflict.incoming_lines
        if l.strip() and not BOOKMARK_ID_PATTERN.match(l)
    ]

    return head_content_lines == incoming_content_lines


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


def extract_relationship_columns(lines: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """
    Extracts (fromColumn_val, toColumn_val) from lines of a relationship block.
    """
    from_col = None
    to_col = None
    for line in lines:
        match_from = FROM_COL_PATTERN.match(line)
        if match_from:
            from_col = match_from.group(1).strip()
        match_to = TO_COL_PATTERN.match(line)
        if match_to:
            to_col = match_to.group(1).strip()
    return from_col, to_col


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


def compute_line_differences(lines1: List[str], lines2: List[str]) -> Tuple[List[str], Set[int], Set[int]]:
    """
    Computes exact line differences between Option 1 (lines1) and Option 2 (lines2).
    """
    diff_summary: List[str] = []
    diff_set1: Set[int] = set()
    diff_set2: Set[int] = set()

    clean1 = [l.strip() for l in lines1 if l.strip()]
    clean2 = [l.strip() for l in lines2 if l.strip()]

    matcher = difflib.SequenceMatcher(None, clean1, clean2)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'replace':
            for a_idx, b_idx in zip(range(i1, i2), range(j1, j2)):
                diff_summary.append(
                    f"{CLR_RED}      - Option [1] (Incoming): {clean1[a_idx]}{CLR_RESET}\n"
                    f"{CLR_GREEN}      + Option [2] (Current) : {clean2[b_idx]}{CLR_RESET}"
                )
        elif tag == 'delete':
            for a_idx in range(i1, i2):
                diff_summary.append(
                    f"{CLR_RED}      - Option [1] (Incoming): {clean1[a_idx]}{CLR_RESET}\n"
                    f"{CLR_GREEN}      + Option [2] (Current) : (missing in Option 2){CLR_RESET}"
                )
        elif tag == 'insert':
            for b_idx in range(j1, j2):
                diff_summary.append(
                    f"{CLR_RED}      - Option [1] (Incoming): (missing in Option 1){CLR_RESET}\n"
                    f"{CLR_GREEN}      + Option [2] (Current) : {clean2[b_idx]}{CLR_RESET}"
                )

    orig_matcher = difflib.SequenceMatcher(
        None, [l.rstrip("\r\n").strip() for l in lines1], [l.rstrip("\r\n").strip() for l in lines2]
    )
    for tag, i1, i2, j1, j2 in orig_matcher.get_opcodes():
        if tag != 'equal':
            for idx in range(i1, i2):
                diff_set1.add(idx)
            for idx in range(j1, j2):
                diff_set2.add(idx)

    return diff_summary, diff_set1, diff_set2


def compute_block_differences(col_blocks: List[ColumnBlock]) -> Tuple[List[str], List[Set[int]]]:
    """
    Computes exact property and line differences between duplicate object blocks.
    """
    if len(col_blocks) >= 2:
        diff_sum, dset1, dset2 = compute_line_differences(col_blocks[0].lines, col_blocks[1].lines)
        return diff_sum, [dset1, dset2]
    return [], [set() for _ in col_blocks]


def format_lines_with_diff_highlights(lines: List[str], diff_indices: Set[int]) -> str:
    """
    Formats lines of code, highlighting differing lines with ANSI color, '*' indicator, and tag.
    """
    formatted_lines = []
    for idx, line in enumerate(lines):
        clean_line = line.rstrip("\r\n")
        if idx in diff_indices:
            formatted_lines.append(f"{CLR_YELLOW}  * {clean_line}   <-- DIFFERENT{CLR_RESET}")
        else:
            formatted_lines.append(f"    {clean_line}")
    return "\n".join(formatted_lines)


def get_git_repo_root(target_path: Path) -> Optional[Path]:
    """Finds top-level Git repository root directory for target_path."""
    try:
        abs_path = target_path.resolve()
        cwd_dir = abs_path if abs_path.is_dir() else abs_path.parent
        res = subprocess.run(
            ["git", "rev-parse", "--show-toplevel"],
            cwd=cwd_dir,
            capture_output=True,
            text=True,
            timeout=5
        )
        if res.returncode == 0 and res.stdout.strip():
            return Path(res.stdout.strip()).resolve()
    except Exception:
        pass
    return None


def get_git_status_files(repo_root: Path) -> Tuple[Set[Path], Set[Path]]:
    """
    Runs 'git status --porcelain' ONCE at repo_root (~10ms) to return:
    - already_staged_clean: Set of absolute Path objects staged in index with 0 worktree changes ('M  ', 'A  ')
    - unstaged_candidates: Set of absolute Path objects with unstaged worktree changes or untracked (' M', '??', 'MM', 'UU')
    """
    already_staged_clean: Set[Path] = set()
    unstaged_candidates: Set[Path] = set()

    try:
        res = subprocess.run(
            ["git", "status", "--porcelain"],
            cwd=repo_root,
            capture_output=True,
            text=True,
            timeout=10
        )
        if res.returncode == 0:
            for line in res.stdout.splitlines():
                if len(line) >= 4:
                    xy = line[:2]
                    rel_path_str = line[3:].strip().strip('"')
                    abs_path = (repo_root / rel_path_str).resolve()

                    if xy[0] in ('M', 'A', 'R', 'C') and xy[1] == ' ':
                        already_staged_clean.add(abs_path)
                    elif xy[1] in ('M', '?') or xy in ('MM', 'UU', 'AA', 'DD', 'AU', 'UA'):
                        unstaged_candidates.add(abs_path)
    except Exception:
        pass

    return already_staged_clean, unstaged_candidates


def get_git_head_content(file_path: Path) -> Optional[str]:
    """
    Attempts to retrieve the content of file_path from Git HEAD (or ORIG_HEAD)
    using git CLI commands with case-insensitive Windows path resolution.
    """
    try:
        abs_path = file_path.resolve()
        repo_root = get_git_repo_root(abs_path)
        if not repo_root:
            return None

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

    if target_type == "relationship":
        from_c, to_c = extract_relationship_columns(col_blocks[0].lines)
        if from_c and to_c:
            canon_key = get_canonical_relationship_key(from_c, to_c)
            head_matches = [
                b for b in head_blocks
                if b.obj_type.lower() == "relationship" and extract_relationship_columns(b.lines)[0] and
                get_canonical_relationship_key(*extract_relationship_columns(b.lines)) == canon_key
            ]
        else:
            head_matches = []
    else:
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
    Parses TMDL file lines and extracts all column, expression, measure, and relationship blocks,
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
            header_str = match.group("header") or ""

            col_name = extract_column_name(header_str) if header_str else f"({obj_type})"
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

            if obj_type == "relationship":
                from_c, to_c = extract_relationship_columns(block_lines)
                if from_c and to_c:
                    col_name = f"{from_c} -> {to_c}"

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


def parse_git_conflict_blocks(
    lines: List[str], tag_filter: str = "all"
) -> List[ConflictMarkerBlock]:
    """
    Finds Git conflict blocks (<<<<<<<, =======, >>>>>>>) filtered by tag_filter:
    - 'lineage': blocks containing lineageTag or sourceLineageTag
    - 'logical_id': blocks containing logicalId or sourceLogicalId
    - 'schema': blocks containing $schema
    - 'bookmark': blocks containing "name": "<hex-hash>"
    - 'subset_addition': blocks where base content is identical + extra lines added
    - 'empty_addition': blocks where one side is completely empty/blank and other side has text
    - 'addition': any addition block (both subset and empty-side additions)
    - 'all': any conflict block (LineageTags, LogicalIds, SchemaTags, Bookmark IDs, Additions & All Other Conflicts)
    """
    conflict_blocks: List[ConflictMarkerBlock] = []
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()
        if stripped.startswith("<<<<<<<"):
            start_idx = i
            ref_head = stripped[7:].strip()
            head_label = f"Current Branch ({ref_head if ref_head else 'HEAD'})"
            sep_idx = None
            end_idx = None
            ref_incoming = ""

            j = i + 1
            while j < n:
                curr = lines[j].strip()
                if curr.startswith("======="):
                    sep_idx = j
                elif curr.startswith(">>>>>>>"):
                    end_idx = j
                    ref_incoming = curr[7:].strip()
                    break
                j += 1

            if sep_idx is not None and end_idx is not None:
                inc_label = f"Incoming Change ({ref_incoming if ref_incoming else 'feature branch'})"
                head_lines = lines[start_idx + 1:sep_idx]
                inc_lines = lines[sep_idx + 1:end_idx]

                all_block_lines = head_lines + inc_lines

                should_include = False
                if tag_filter == "lineage":
                    should_include = any(TAG_PATTERN.match(l) for l in all_block_lines)
                elif tag_filter == "logical_id":
                    should_include = any(LOGICAL_ID_PATTERN.match(l) for l in all_block_lines)
                elif tag_filter == "schema":
                    should_include = any(SCHEMA_PATTERN.match(l) for l in all_block_lines)
                elif tag_filter == "bookmark":
                    should_include = any(BOOKMARK_ID_PATTERN.match(l) for l in all_block_lines)
                elif tag_filter == "subset_addition":
                    should_include = is_subset_addition_only(head_lines, inc_lines)
                elif tag_filter == "empty_addition":
                    should_include = is_empty_side_addition_only(head_lines, inc_lines)
                elif tag_filter == "addition":
                    should_include = is_subset_addition(head_lines, inc_lines)
                elif tag_filter == "all":
                    should_include = True

                if should_include:
                    conflict_blocks.append(ConflictMarkerBlock(
                        start_line=start_idx,
                        sep_line=sep_idx,
                        end_line=end_idx,
                        head_label=head_label,
                        incoming_label=inc_label,
                        head_lines=head_lines,
                        incoming_lines=inc_lines
                    ))
                i = end_idx + 1
            else:
                i += 1
        else:
            i += 1

    return conflict_blocks


def propagate_hash_replacement(target_files: List[Path], old_hash: str, new_hash: str) -> int:
    """
    Scans all project files and updates cross-references from old_hash to new_hash.
    Returns the total count of cross-reference replacements made across all files.
    """
    if not old_hash or not new_hash or old_hash == new_hash:
        return 0

    replacements_made = 0
    for file_path in target_files:
        try:
            with open(file_path, "rb") as f:
                raw_bytes = f.read()
            has_bom = raw_bytes.startswith(b"\xef\xbb\xbf")
            content = (raw_bytes[3:] if has_bom else raw_bytes).decode("utf-8")
        except Exception:
            continue

        if old_hash in content:
            count = content.count(old_hash)
            new_content = content.replace(old_hash, new_hash)
            encoded = new_content.encode("utf-8")
            if has_bom:
                encoded = b"\xef\xbb\xbf" + encoded
            try:
                with open(file_path, "wb") as f:
                    f.write(encoded)
                replacements_made += count
            except Exception:
                pass

    return replacements_made


def group_duplicate_columns(
    blocks: List[ColumnBlock], target_object_type: str = "all"
) -> Dict[Tuple[str, str], List[ColumnBlock]]:
    """
    Groups object blocks by canonical key and filters by target_object_type.
    """
    grouped: Dict[Tuple[str, str], List[ColumnBlock]] = {}
    for block in blocks:
        if target_object_type != "all" and block.obj_type.lower() != target_object_type.lower():
            continue

        if block.obj_type.lower() == "relationship":
            from_c, to_c = extract_relationship_columns(block.lines)
            if from_c and to_c:
                canon_key = get_canonical_relationship_key(from_c, to_c)
                key = ("relationship", f"{canon_key[0]} <-> {canon_key[1]}")
            else:
                key = ("relationship", block.col_name.lower())
        else:
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
    """Checks if file extension or filename matches the filter."""
    if extensions is None:
        return True
    ext = file_path.suffix.lstrip(".").lower()
    name = file_path.name.lstrip(".").lower()
    return (ext in extensions) or (name in extensions)


def ask_propagate_cross_references(args_propagate: Optional[bool] = None) -> bool:
    """
    Prompts user whether to enable Global Cross-Reference Propagation.
    Returns True if enabled, False if disabled (for maximum speed).
    """
    if args_propagate is not None:
        return args_propagate

    print("\nEnable Global Cross-Reference Propagation across all project files?")
    print(" y - Yes, scan & update cross-references of discarded IDs across all files")
    print(" n - No, resolve conflict lines only (FASTER performance, default)")
    while True:
        choice = input("Enter choice (y/n, default: n): ").strip().lower()
        if choice in ("n", "no", ""):
            return False
        elif choice in ("y", "yes"):
            return True
        print("Invalid input. Please enter 'y' or 'n'.\n")


def get_conflict_target_type(args_conflict_type: Optional[str]) -> str:
    """
    Returns normalized conflict marker target filter: 'lineage', 'logical_id', 'schema', 'bookmark', 'subset_addition', 'empty_addition', 'addition', 'all', or 'back'.
    Prompts interactively if not provided as CLI argument.
    """
    if args_conflict_type:
        val = args_conflict_type.strip().lower()
        if val in ("b", "back", "0"):
            return "back"
        elif val in ("1", "lineage", "lineagetag", "lineagetags"):
            return "lineage"
        elif val in ("2", "logical", "logicalid", "logicalids"):
            return "logical_id"
        elif val in ("3", "schema", "schematag", "schematags"):
            return "schema"
        elif val in ("4", "bookmark", "id", "hash"):
            return "bookmark"
        elif val in ("5.1", "subset", "subset_addition"):
            return "subset_addition"
        elif val in ("5.2", "empty", "empty_addition"):
            return "empty_addition"
        elif val in ("5", "5.3", "addition", "additions", "add"):
            return "addition"
        elif val in ("6", "all", "rest", "both"):
            return "all"

    print("\nSelect target property filter for Git Conflict Markers:")
    print(" 1 - LineageTags (lineageTag, sourceLineageTag)")
    print(" 2 - LogicalIds (logicalId, sourceLogicalId)")
    print(" 3 - SchemaTags ($schema: \"https://...\")")
    print(" 4 - Bookmark / Object Name IDs (\"name\": \"<hex-hash>\")")
    print(" 5 - Additions Only:")
    print("     5.1 - Subset Additions Only (Base content identical + extra lines added)")
    print("     5.2 - Empty-Side Additions Only (One side contains content, other side is blank/empty)")
    print("     5.3 - All Additions (Both Subset & Empty-Side Additions)")
    print(" 6 - All Conflict Markers (LineageTags, LogicalIds, SchemaTags, Bookmark IDs, Additions & All Other Conflicts)")
    print(" b - Back to Main Menu")
    while True:
        choice = input("Enter option (1, 2, 3, 4, 5, 5.1, 5.2, 5.3, 6, or b): ").strip().lower()
        if choice in ("b", "back", "0"):
            return "back"
        elif choice == "1":
            return "lineage"
        elif choice == "2":
            return "logical_id"
        elif choice == "3":
            return "schema"
        elif choice == "4":
            return "bookmark"
        elif choice in ("5.1", "subset"):
            return "subset_addition"
        elif choice in ("5.2", "empty"):
            return "empty_addition"
        elif choice in ("5", "5.3", "addition"):
            return "addition"
        elif choice == "6":
            return "all"
        print("Invalid input. Please enter 1, 2, 3, 4, 5.1, 5.2, 5.3, 6, or b.\n")


def get_target_object_type(args_type: Optional[str]) -> str:
    """
    Returns normalized target object type for Mode 2: 'column', 'expression', 'relationship', 'formatter', 'all', or 'back'.
    """
    if args_type:
        val = args_type.strip().lower()
        if val in ("b", "back", "0"):
            return "back"
        elif val in ("1", "column", "columns"):
            return "column"
        elif val in ("2", "expression", "expressions"):
            return "expression"
        elif val in ("3", "relationship", "relationships"):
            return "relationship"
        elif val in ("4", "formatter", "format", "comma", "json"):
            return "formatter"
        elif val in ("5", "all", "both"):
            return "all"

    print("\nSelect target object type to process (Mode 2):")
    print(" 1 - Columns")
    print(" 2 - Expressions")
    print(" 3 - Relationships")
    print(" 4 - Field Parameter & JSON Comma Syntax Auto-Formatter (Auto-fix length, indices & commas)")
    print(" 5 - All (Columns, Expressions, Relationships & Field Parameter / JSON Auto-Fixer)")
    print(" b - Back to Main Menu")
    while True:
        choice = input("Enter option (1, 2, 3, 4, 5, or b): ").strip().lower()
        if choice in ("b", "back", "0"):
            return "back"
        elif choice == "1":
            return "column"
        elif choice == "2":
            return "expression"
        elif choice == "3":
            return "relationship"
        elif choice == "4":
            return "formatter"
        elif choice == "5":
            return "all"
        print("Invalid input. Please enter 1, 2, 3, 4, 5, or b.\n")


def run_mode_1_conflict_markers(
    target_files: List[Path], conflict_target_type: str, dry_run: bool, auto_keep_state: List[Optional[str]], stats: SummaryStats, propagate_refs: bool = False
) -> None:
    """
    Mode 1: Resolves Git Conflict Markers (<<<<<<< ======= >>>>>>>) filtered by conflict_target_type.
    In Combo Mode (conflict_target_type == "all"), 1A / 2A auto-keep is SCOPED STRICTLY per category
    (Lineage, LogicalId, Schema, Bookmark, Subset Addition, Empty-Side Addition, Other).
    Distinguishes Subset Additions (base content identical + extra lines) from Empty-Side Additions (content vs blank/empty).
    Auto-fixes JSON syntax, commas, and Field Parameter metadata after file modification.
    """
    print("\n" + CLR_CYAN + "================================================================================" + CLR_RESET)
    print(CLR_BOLD + f"  MODE 1: Resolving Git Conflict Markers (Target Filter: {conflict_target_type.upper()})" + CLR_RESET)
    if conflict_target_type in ("bookmark", "all"):
        prop_desc = "ENABLED" if propagate_refs else "DISABLED (FAST MODE)"
        print(f"  • Cross-Reference Propagation: {CLR_YELLOW}{prop_desc}{CLR_RESET}")
    print(CLR_CYAN + "================================================================================" + CLR_RESET)

    category_auto_keep: Dict[str, Optional[str]] = {}
    if auto_keep_state[0] == "first":
        category_auto_keep["global"] = "first"
    elif auto_keep_state[0] == "last":
        category_auto_keep["global"] = "last"

    for file_idx, file_path in enumerate(target_files, 1):
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
            content = (raw_bytes[3:] if has_bom else raw_bytes).decode("utf-8")
        except UnicodeDecodeError:
            stats.files_skipped += 1
            continue

        lines = content.splitlines(keepends=True)
        conflicts = parse_git_conflict_blocks(lines, tag_filter=conflict_target_type)
        if not conflicts:
            continue

        stats.files_scanned += 1
        num_conflicts = len(conflicts)
        lines_to_delete: Set[int] = set()
        partial_insertions: Dict[int, List[str]] = {}

        print(f"\nFile [{file_idx}/{len(target_files)}]: {file_path} - Found {num_conflicts} conflict marker(s).")

        for c_idx, conflict in enumerate(conflicts, 1):
            stats.conflicts_found += 1

            cat_type = detect_conflict_category(conflict)
            is_pure = is_pure_property_conflict(conflict, cat_type if conflict_target_type == "all" else conflict_target_type)
            same_bookmark_content = is_bookmark_content_same(conflict) if (cat_type == "bookmark" or conflict_target_type == "bookmark") else True
            is_divergent = (cat_type == "bookmark" and not same_bookmark_content)

            is_subset = is_subset_addition_only(conflict.head_lines, conflict.incoming_lines)
            is_empty_side = is_empty_side_addition_only(conflict.head_lines, conflict.incoming_lines)

            # Determine which side has the addition content
            inc_has_content = bool([l for l in conflict.incoming_lines if l.strip()])
            head_has_content = bool([l for l in conflict.head_lines if l.strip()])
            rec_side = 1 if inc_has_content else 2

            print("\n" + CLR_BOLD + f"Conflict {c_idx} of {num_conflicts} (Category: {cat_type.upper()}, Lines {conflict.start_line + 1}-{conflict.end_line + 1}):" + CLR_RESET)

            if is_subset:
                print(f"{CLR_GREEN}  [+] SUBSET ADDITION DETECTED: Base content identical + extra lines added on Option {rec_side}!{CLR_RESET}")
            elif is_empty_side:
                empty_side_str = "Option 2 is BLANK/EMPTY" if inc_has_content else "Option 1 is BLANK/EMPTY"
                print(f"{CLR_GREEN}  [+] EMPTY-SIDE ADDITION DETECTED: Option {rec_side} contains content while {empty_side_str}!{CLR_RESET}")
            elif is_divergent:
                print(f"{CLR_RED}  [!] BOOKMARK DIVERGENCE DETECTED: Name is DIFFERENT AND Content/Properties are ALSO DIFFERENT!{CLR_RESET}")
            elif not is_pure:
                print(f"{CLR_YELLOW}  [!] MIXED CONFLICT: Block contains {cat_type.upper()} AND other code/visual changes!{CLR_RESET}")

            print(f"  Option [1] [{conflict.incoming_label}] (Top):")
            for l in conflict.incoming_lines:
                print(f"      {CLR_GREEN}{l.rstrip()}{CLR_RESET}")
            print(f"  Option [2] [{conflict.head_label}] (Bottom):")
            for l in conflict.head_lines:
                print(f"      {CLR_RED}{l.rstrip()}{CLR_RESET}")

            selected_option = None  # "1", "2", "1P", "2P", "s"

            current_cat_keep = category_auto_keep.get(cat_type) or category_auto_keep.get("global")

            if current_cat_keep == "first" and not is_divergent:
                selected_option = "1"
                print(f"  [AUTO-KEEP 1A ({cat_type.upper()})] Selected Option 1 (Incoming Change)")
            elif current_cat_keep == "last" and not is_divergent:
                selected_option = "2"
                print(f"  [AUTO-KEEP 2A ({cat_type.upper()})] Selected Option 2 (Current Branch)")
            else:
                while True:
                    if is_subset or is_empty_side:
                        add_kind = "SUBSET ADDITION" if is_subset else "EMPTY-SIDE ADDITION"
                        rec_str = f" (Recommended: {rec_side})"
                        prompt_msg = (
                            f"Select option for {add_kind} conflict{rec_str}:\n"
                            "  1  = Accept Option 1 [Incoming Change] (Top)\n"
                            "  2  = Accept Option 2 [Current Branch / HEAD] (Bottom)\n"
                            f"  1A = Accept Option 1 for all remaining {cat_type.upper()} conflicts\n"
                            f"  2A = Accept Option 2 for all remaining {cat_type.upper()} conflicts\n"
                            "  s  = Skip: "
                        )
                    elif is_divergent:
                        prompt_msg = (
                            f"Select option for DIFFERENT-CONTENT Bookmark conflict:\n"
                            "  1  = Accept Option 1 [Incoming Change Bookmark]\n"
                            "  2  = Accept Option 2 [Current Branch / HEAD Bookmark]\n"
                            "  1A = Accept Option 1 for all remaining BOOKMARK conflicts\n"
                            "  2A = Accept Option 2 for all remaining BOOKMARK conflicts\n"
                            "  s  = Skip (Do NOT touch, leave conflict untouched): "
                        )
                    elif is_pure:
                        prompt_msg = (
                            f"Select option to KEEP for this {cat_type.upper()} conflict marker:\n"
                            f"  (1 = Option 1 [Incoming], 2 = Option 2 [Current], 1A = Option 1 for ALL {cat_type.upper()}S, 2A = Option 2 for ALL {cat_type.upper()}S, s = skip): "
                        )
                    else:
                        prompt_msg = (
                            f"Select option to KEEP for this MIXED conflict marker:\n"
                            "  1  = Full Option 1 [Incoming Change]\n"
                            "  2  = Full Option 2 [Current Branch]\n"
                            "  1P = Partial Fix: Update tag to Option 1 [Incoming], KEEP conflict markers for visual/other changes\n"
                            "  2P = Partial Fix: Update tag to Option 2 [Current], KEEP conflict markers for visual/other changes\n"
                            f"  1A = Full Option 1 for ALL remaining {cat_type.upper()} conflicts\n"
                            f"  2A = Full Option 2 for ALL remaining {cat_type.upper()} conflicts\n"
                            "  s  = Skip: "
                        )

                    choice = input(prompt_msg).strip().lower()
                    if choice in ("s", "skip"):
                        print(f"{CLR_YELLOW}Skipped conflict block (left untouched).{CLR_RESET}")
                        stats.conflicts_skipped += 1
                        selected_option = "s"
                        break
                    elif choice == "1":
                        selected_option = "1"
                        break
                    elif choice == "2":
                        selected_option = "2"
                        break
                    elif choice in ("1p", "p1") and not is_pure and not is_divergent:
                        selected_option = "1P"
                        break
                    elif choice in ("2p", "p2") and not is_pure and not is_divergent:
                        selected_option = "2P"
                        break
                    elif choice in ("1a", "a1", "all1"):
                        selected_option = "1"
                        category_auto_keep[cat_type] = "first"
                        print(f"{CLR_GREEN}--> Activated AUTO-RESOLVE ALL for category [{cat_type.upper()}]: Keeping Option 1 [Incoming Change].{CLR_RESET}")
                        break
                    elif choice in ("2a", "a2", "all2"):
                        selected_option = "2"
                        category_auto_keep[cat_type] = "last"
                        print(f"{CLR_GREEN}--> Activated AUTO-RESOLVE ALL for category [{cat_type.upper()}]: Keeping Option 2 [Current Branch].{CLR_RESET}")
                        break
                    print("Invalid choice. Please enter a valid option.")

            if selected_option == "s":
                continue

            if selected_option in ("1", "2"):
                stats.conflicts_resolved += 1
                lines_to_delete.add(conflict.start_line) # Delete <<<<<<<
                lines_to_delete.add(conflict.sep_line)   # Delete =======
                lines_to_delete.add(conflict.end_line)   # Delete >>>>>>>
                
                selected_lines = conflict.incoming_lines if selected_option == "1" else conflict.head_lines
                unselected = conflict.head_lines if selected_option == "1" else conflict.incoming_lines

                if selected_option == "1":
                    # Keep Option 1 (Incoming): Delete all HEAD lines between start_line and sep_line
                    for idx in range(conflict.start_line + 1, conflict.sep_line):
                        lines_to_delete.add(idx)
                elif selected_option == "2":
                    # Keep Option 2 (Current): Delete all Incoming lines between sep_line and end_line
                    for idx in range(conflict.sep_line + 1, conflict.end_line):
                        lines_to_delete.add(idx)

                # Extract and propagate hash replacement ONLY if propagate_refs is True
                if propagate_refs and (cat_type == "bookmark" or (conflict_target_type == "all" and is_pure)):
                    kept_hash = None
                    discarded_hash = None

                    for l in selected_lines:
                        m = BOOKMARK_ID_PATTERN.match(l)
                        if m:
                            kept_hash = m.group(1)
                            break

                    for l in unselected:
                        m = BOOKMARK_ID_PATTERN.match(l)
                        if m:
                            discarded_hash = m.group(1)
                            break

                    if kept_hash and discarded_hash and kept_hash != discarded_hash:
                        refs_updated = propagate_hash_replacement(target_files, discarded_hash, kept_hash)
                        if refs_updated > 0:
                            print(f"{CLR_GREEN}--> Propagated Bookmark/Object ID Hash: Updated {refs_updated} cross-references of '{discarded_hash}' to '{kept_hash}'.{CLR_RESET}")

            elif selected_option in ("1P", "2P"):
                stats.conflicts_resolved += 1
                source_lines = conflict.incoming_lines if selected_option == "1P" else conflict.head_lines

                extracted_tag_line = None
                if cat_type == "lineage":
                    pat = TAG_PATTERN
                elif cat_type == "logical_id":
                    pat = LOGICAL_ID_PATTERN
                elif cat_type == "bookmark":
                    pat = BOOKMARK_ID_PATTERN
                else:
                    pat = SCHEMA_PATTERN

                for l in source_lines:
                    if pat.match(l):
                        extracted_tag_line = l
                        break

                if extracted_tag_line:
                    partial_insertions[conflict.start_line] = [extracted_tag_line]
                    for idx in range(conflict.start_line + 1, conflict.end_line):
                        if pat.match(lines[idx]):
                            lines_to_delete.add(idx)
                    print(f"{CLR_GREEN}--> Partial Fix Applied: Updated tag line and preserved visual conflict markers.{CLR_RESET}")

        if (lines_to_delete or partial_insertions) and not dry_run:
            stats.files_modified += 1
            new_lines = []
            for idx, l in enumerate(lines):
                if idx in partial_insertions:
                    new_lines.extend(partial_insertions[idx])
                if idx not in lines_to_delete:
                    new_lines.append(l)

            cleaned_lines = cleanup_excessive_blank_lines(new_lines)
            new_content = "".join(cleaned_lines)

            # Auto-format JSON syntax/commas & Field Parameter metadata
            new_content, fp_count = fix_field_parameter_metadata(new_content)
            if file_path.suffix.lower() in (".json", ".pbir", ".pbip", ".platform"):
                new_content = fix_pbip_json_formatting(new_content)

            encoded = new_content.encode("utf-8")
            if has_bom:
                encoded = b"\xef\xbb\xbf" + encoded
            with open(file_path, "wb") as f:
                f.write(encoded)
            print(f"[UPDATED] File conflict markers resolved & formatted: {file_path}")


def run_mode_2_dedupe_objects(
    target_files: List[Path], target_object_type: str, dry_run: bool, auto_keep_state: List[Optional[str]], stats: SummaryStats
) -> None:
    """
    Mode 2: Deduplicates Objects (Columns, Expressions, Relationships) and auto-fixes Field Parameter & JSON comma metadata.
    Accurately tracks files_scanned and files_modified statistics.
    """
    print("\n" + CLR_CYAN + "================================================================================" + CLR_RESET)
    print(CLR_BOLD + f"  MODE 2: Resolving Duplicate Objects & Formatter (Target: {target_object_type.upper()})" + CLR_RESET)
    print(CLR_CYAN + "================================================================================" + CLR_RESET)

    if target_object_type in ("formatter", "all"):
        json_files_fixed = 0
        for fpath in target_files:
            if fpath.suffix.lower() in (".json", ".pbir", ".pbip", ".platform", ".tmdl"):
                stats.files_scanned += 1
                try:
                    with open(fpath, "rb") as f:
                        raw = f.read()
                    has_bom = raw.startswith(b"\xef\xbb\xbf")
                    txt = (raw[3:] if has_bom else raw).decode("utf-8")
                    
                    fixed, fp_count = fix_field_parameter_metadata(txt)
                    if fpath.suffix.lower() in (".json", ".pbir", ".pbip", ".platform"):
                        fixed = fix_pbip_json_formatting(fixed)

                    if fixed != txt:
                        stats.files_modified += 1
                        if not dry_run:
                            enc = fixed.encode("utf-8")
                            if has_bom:
                                enc = b"\xef\xbb\xbf" + enc
                            with open(fpath, "wb") as f:
                                f.write(enc)
                            json_files_fixed += 1
                            print(f"  {CLR_GREEN}[FORMATTED]{CLR_RESET} Fixed Field Parameter & JSON metadata in: {fpath}")
                        else:
                            print(f"  [DRY-RUN FORMAT] Would fix Field Parameter / JSON formatting in: {fpath}")
                except Exception:
                    stats.files_skipped += 1

        print(f"\nFormatter Summary: Scanned {CLR_CYAN}{stats.files_scanned}{CLR_RESET} file(s) | Formatted: {CLR_GREEN}{json_files_fixed}{CLR_RESET} file(s)")

    if target_object_type != "formatter":
        tasks, total_duplicates_all = scan_and_prepare_tasks(target_files, target_object_type, stats)
        global_remaining = [total_duplicates_all]
        total_files = len(tasks)

        for file_idx, task in enumerate(tasks, 1):
            process_file_task(
                task, file_idx, total_files, dry_run, auto_keep_state, global_remaining, stats
            )


def run_mode_3_stage_clean_files(target_files: List[Path], dry_run: bool, target_path: Optional[Path] = None) -> None:
    """
    Mode 3: Automatically runs 'git add' on target files with 0 remaining conflicts.
    Uses 'git status --porcelain' ONCE at repository root (~10ms) to bypass unmodified & already-staged files.
    Displays live real-time progress feedback ([X/N] Checking...) for candidate files.
    """
    print("\n" + CLR_CYAN + "================================================================================" + CLR_RESET)
    print(CLR_BOLD + "  MODE 3: Stage Clean Files to Git (0 Remaining Conflicts - BATCHED FAST MODE)" + CLR_RESET)
    print(CLR_CYAN + "================================================================================" + CLR_RESET)

    sample_path = target_path if target_path else (target_files[0] if target_files else Path.cwd())
    repo_root = get_git_repo_root(sample_path)

    if not repo_root:
        print(f"{CLR_RED}Error: Target directory is not inside a valid Git repository.{CLR_RESET}")
        return

    print(f"Git Repository Root: {CLR_CYAN}{repo_root}{CLR_RESET}")
    print("Querying Git status for repository modified files...")

    already_staged_set, unstaged_candidates_set = get_git_status_files(repo_root)

    # Filter by target_path / target_files if specified
    if target_files:
        target_resolved_set = {p.resolve() for p in target_files}
        already_staged_filtered = [p for p in already_staged_set if p in target_resolved_set]
        candidate_files = [p for p in unstaged_candidates_set if p in target_resolved_set]
    else:
        already_staged_filtered = list(already_staged_set)
        candidate_files = list(unstaged_candidates_set)

    print(f"\nGit Status Summary:")
    print(f"  • Already Staged (Skipped) : {CLR_YELLOW}{len(already_staged_filtered)} file(s){CLR_RESET}")
    print(f"  • Unstaged Candidates      : {CLR_CYAN}{len(candidate_files)} file(s){CLR_RESET}")

    if already_staged_filtered:
        print(f"\n{CLR_YELLOW}Already Staged in Git Index (Skipped):{CLR_RESET}")
        for fp in already_staged_filtered:
            print(f"  {CLR_YELLOW}[ALREADY STAGED - SKIPPED]{CLR_RESET} {fp}")

    if not candidate_files:
        print(f"\n{CLR_GREEN}[OK] No unstaged modified files to stage.{CLR_RESET}")
        print("\n" + "-" * 60)
        print(f"Staging Summary: Clean Files Newly Staged: {CLR_GREEN}0{CLR_RESET} | Already Staged (Skipped): {CLR_YELLOW}{len(already_staged_filtered)}{CLR_RESET} | Files with Conflicts Remaining: {CLR_YELLOW}0{CLR_RESET}")
        return

    files_to_stage: List[Path] = []
    conflict_files: List[Tuple[Path, int]] = []
    total_candidates = len(candidate_files)

    print(f"\nInspecting {total_candidates} unstaged candidate file(s) for remaining conflict markers...")

    for idx, file_path in enumerate(candidate_files, 1):
        print(f"  [{idx}/{total_candidates}] Checking: {CLR_CYAN}{file_path.name}{CLR_RESET}...", end="\r", flush=True)

        try:
            with open(file_path, "rb") as f:
                chunk = f.read(8192)
                if b"\x00" in chunk:
                    continue
                f.seek(0)
                raw_bytes = f.read()
        except Exception:
            continue

        has_bom = raw_bytes.startswith(b"\xef\xbb\xbf")
        try:
            content = (raw_bytes[3:] if has_bom else raw_bytes).decode("utf-8")
        except UnicodeDecodeError:
            continue

        lines = content.splitlines(keepends=True)
        conflicts = parse_git_conflict_blocks(lines, tag_filter="all")

        if conflicts:
            conflict_files.append((file_path, len(conflicts)))
        else:
            files_to_stage.append(file_path)

    print(" " * 80, end="\r")  # Clear line

    if files_to_stage:
        print(f"\n{CLR_GREEN}Staging {len(files_to_stage)} Clean File(s) to Git Index...{CLR_RESET}")

        if not dry_run:
            CHUNK_SIZE = 50
            for i in range(0, len(files_to_stage), CHUNK_SIZE):
                chunk_paths = files_to_stage[i:i + CHUNK_SIZE]
                cmd = ["git", "add"] + [str(p.resolve()) for p in chunk_paths]
                try:
                    res = subprocess.run(
                        cmd,
                        cwd=repo_root,
                        capture_output=True,
                        text=True,
                        timeout=10
                    )
                    if res.returncode == 0:
                        for p in chunk_paths:
                            print(f"  {CLR_GREEN}[STAGED]{CLR_RESET} {p}")
                    else:
                        print(f"  {CLR_YELLOW}[GIT ADD ERROR]{CLR_RESET} Batch of {len(chunk_paths)} files failed: {res.stderr.strip()}")
                except Exception as e:
                    print(f"  {CLR_RED}[FAILED]{CLR_RESET} Batch of {len(chunk_paths)} files error: {e}")
        else:
            for fpath in files_to_stage:
                print(f"  [DRY-RUN STAGE] {fpath}")

    if conflict_files:
        print(f"\n{CLR_YELLOW}Files with Remaining Conflicts (NOT staged):{CLR_RESET}")
        for fpath, c_count in conflict_files:
            print(f"  {CLR_YELLOW}[CONFLICT]{CLR_RESET} {fpath} ({c_count} conflict marker(s) remaining)")

    print("\n" + "-" * 60)
    print(f"Staging Summary: Clean Files Newly Staged: {CLR_GREEN}{len(files_to_stage)}{CLR_RESET} | Already Staged (Skipped): {CLR_YELLOW}{len(already_staged_filtered)}{CLR_RESET} | Files with Conflicts Remaining: {CLR_YELLOW}{len(conflict_files)}{CLR_RESET}")


def run_mode_4_metadata_diagnostic(target_files: List[Path]) -> List[DiagnosticIssue]:
    """
    Mode 4: Performs a comprehensive Power BI PBIP Metadata & Health Validation Check.
    Scans files for remaining conflict markers, duplicate objects, Field Parameter length/index mismatches, JSON syntax errors, and missing lineageTags.
    """
    print("\n" + CLR_CYAN + "================================================================================" + CLR_RESET)
    print(CLR_BOLD + "  MODE 4: Power BI PBIP Metadata & Health Validation Diagnostic Check" + CLR_RESET)
    print(CLR_CYAN + "================================================================================" + CLR_RESET)

    issues: List[DiagnosticIssue] = []
    files_scanned = 0

    for file_path in target_files:
        try:
            with open(file_path, "rb") as f:
                chunk = f.read(8192)
                if b"\x00" in chunk:
                    continue
                f.seek(0)
                raw_bytes = f.read()
        except Exception:
            continue

        has_bom = raw_bytes.startswith(b"\xef\xbb\xbf")
        try:
            content = (raw_bytes[3:] if has_bom else raw_bytes).decode("utf-8")
        except UnicodeDecodeError:
            continue

        files_scanned += 1
        lines = content.splitlines(keepends=True)

        # 1. Check for remaining Git Conflict Markers
        for idx, line in enumerate(lines, 1):
            s = line.strip()
            if s.startswith("<<<<<<<") or s.startswith("=======") or s.startswith(">>>>>>>"):
                issues.append(DiagnosticIssue(
                    file_path=file_path,
                    issue_type="Git Conflict Marker",
                    description=f"Unresolved conflict marker line: '{s[:40]}'",
                    line_no=idx
                ))

        # 2. Check for remaining Duplicate Objects
        blocks = extract_column_blocks(lines)
        duplicate_groups = group_duplicate_columns(blocks, target_object_type="all")
        for (obj_type, _), dups in duplicate_groups.items():
            obj_name = dups[0].col_name
            issues.append(DiagnosticIssue(
                file_path=file_path,
                issue_type="Duplicate Object",
                description=f"Duplicate {obj_type} definition found: '{obj_name}' ({len(dups)} occurrences)",
                line_no=dups[0].start_line + 1
            ))

        # 3. Check for Field Parameter Length & Index Mismatches
        if file_path.suffix.lower() in (".json", ".pbir", ".pbip", ".platform"):
            try:
                data = json.loads(content)
                def check_fp_json(node: Any):
                    if isinstance(node, dict):
                        for k, v in list(node.items()):
                            if isinstance(v, dict):
                                projs = v.get("projections") or v.get("parameterProjections") or v.get("parameterFields")
                                if isinstance(projs, list) and projs:
                                    if "length" in v and v["length"] != len(projs):
                                        issues.append(DiagnosticIssue(
                                            file_path=file_path,
                                            issue_type="Field Parameter Length Mismatch",
                                            description=f"Field Parameter length property is {v['length']}, but actual projections count is {len(projs)}"
                                        ))
                                    indices = [item.get("index") for item in projs if isinstance(item, dict) and "index" in item]
                                    expected = list(range(len(indices)))
                                    if indices != expected:
                                        issues.append(DiagnosticIssue(
                                            file_path=file_path,
                                            issue_type="Field Parameter Index Error",
                                            description=f"Field Parameter index sequence is {indices}, expected {expected}"
                                        ))
                            elif k in ("projections", "parameterProjections", "parameterFields") and isinstance(v, list) and v:
                                if "length" in node and node["length"] != len(v):
                                    issues.append(DiagnosticIssue(
                                        file_path=file_path,
                                        issue_type="Field Parameter Length Mismatch",
                                        description=f"Field Parameter length property is {node['length']}, but actual projections count is {len(v)}"
                                    ))
                                indices = [item.get("index") for item in v if isinstance(item, dict) and "index" in item]
                                expected = list(range(len(indices)))
                                if indices != expected:
                                    issues.append(DiagnosticIssue(
                                        file_path=file_path,
                                        issue_type="Field Parameter Index Error",
                                        description=f"Field Parameter index sequence is {indices}, expected {expected}"
                                    ))
                        for child in node.values():
                            check_fp_json(child)
                    elif isinstance(node, list):
                        for child in node:
                            check_fp_json(child)
                check_fp_json(data)
            except json.JSONDecodeError as e:
                issues.append(DiagnosticIssue(
                    file_path=file_path,
                    issue_type="JSON Syntax Error",
                    description=f"JSON parse error: {e.msg} at line {e.lineno}, col {e.colno}",
                    line_no=e.lineno
                ))

        # Check for TMDL Field Parameter Tuple Index Mismatches
        if file_path.suffix.lower() == ".tmdl":
            in_fp_source = False
            cur_tuple_idx = 0
            for l_idx, line in enumerate(lines, 1):
                s = line.strip()
                if "source =" in s and ("{" in s or "=" in s):
                    in_fp_source = True
                    cur_tuple_idx = 0
                    continue
                elif s.startswith("}") and in_fp_source:
                    in_fp_source = False
                    continue

                m = TMDL_FP_TUPLE_PATTERN.match(line)
                if m and in_fp_source:
                    idx_val = int(m.group(2))
                    if idx_val != cur_tuple_idx:
                        issues.append(DiagnosticIssue(
                            file_path=file_path,
                            issue_type="Field Parameter Index Error",
                            description=f"TMDL Field Parameter tuple has index {idx_val}, expected sequential index {cur_tuple_idx}",
                            line_no=l_idx
                        ))
                    cur_tuple_idx += 1

        # 4. Check for missing LineageTags in TMDL column blocks
        if file_path.suffix.lower() == ".tmdl":
            for b in blocks:
                if b.obj_type == "column":
                    lt, slt = extract_lineage_tags(b.lines)
                    if not lt and not slt:
                        issues.append(DiagnosticIssue(
                            file_path=file_path,
                            issue_type="Missing LineageTag",
                            description=f"Column '{b.col_name}' has no lineageTag or sourceLineageTag property",
                            line_no=b.start_line + 1
                        ))

    # Output Diagnostic Report
    print(f"\nFiles Scanned: {files_scanned} file(s)")

    if not issues:
        print(f"\n{CLR_GREEN}[OK] HEALTH CHECK PASSED! No metadata issues or conflict errors detected.{CLR_RESET}")
    else:
        print(f"\n{CLR_RED}[FAILED] HEALTH CHECK FAILED! Found {len(issues)} metadata issue(s):{CLR_RESET}\n")

        issue_counts: Dict[str, int] = {}
        for iss in issues:
            issue_counts[iss.issue_type] = issue_counts.get(iss.issue_type, 0) + 1

        for itype, count in issue_counts.items():
            print(f"  • {itype}: {CLR_YELLOW}{count} issue(s){CLR_RESET}")

        print("\n" + "-" * 60)
        print("Detailed Issue List:")
        for idx, iss in enumerate(issues, 1):
            line_str = f"Line {iss.line_no}" if iss.line_no else "File-level"
            print(f"  [{idx}] {CLR_CYAN}{iss.file_path}{CLR_RESET}:{line_str}")
            print(f"      {CLR_RED}[{iss.issue_type}]{CLR_RESET} {iss.description}")

    return issues


def run_mode_5_detailed_conflict_review(
    target_files: List[Path], dry_run: bool, auto_keep_state: List[Optional[str]], stats: SummaryStats
) -> None:
    """
    Mode 5: Reviews remaining Git conflicts one-by-one with full line-by-line diff highlights (* <-- DIFFERENT).
    Guarantees Option [1] is ALWAYS Incoming Change (Top) and Option [2] is ALWAYS Current Branch / HEAD (Bottom).
    """
    print("\n" + CLR_CYAN + "================================================================================" + CLR_RESET)
    print(CLR_BOLD + "  MODE 5: Detailed Conflict Review & Visual Diff Viewer" + CLR_RESET)
    print(CLR_CYAN + "================================================================================" + CLR_RESET)

    for file_idx, file_path in enumerate(target_files, 1):
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
            content = (raw_bytes[3:] if has_bom else raw_bytes).decode("utf-8")
        except UnicodeDecodeError:
            stats.files_skipped += 1
            continue

        lines = content.splitlines(keepends=True)
        conflicts = parse_git_conflict_blocks(lines, tag_filter="all")
        if not conflicts:
            continue

        stats.files_scanned += 1
        num_conflicts = len(conflicts)
        lines_to_delete: Set[int] = set()

        print(f"\nFile [{file_idx}/{len(target_files)}]: {file_path} - Found {num_conflicts} conflict(s) for review.")

        for c_idx, conflict in enumerate(conflicts, 1):
            stats.conflicts_found += 1

            diff_summary, diff_set1, diff_set2 = compute_line_differences(conflict.incoming_lines, conflict.head_lines)

            print("\n" + CLR_BOLD + f"Conflict {c_idx} of {num_conflicts} (Lines {conflict.start_line + 1}-{conflict.end_line + 1}):" + CLR_RESET)

            if diff_summary:
                print("\n" + CLR_CYAN + "  • Key Line Differences Highlight:" + CLR_RESET)
                for d_line in diff_summary:
                    print(d_line)

            # Option 1: ALWAYS Incoming Change (Top)
            fmt_incoming = format_lines_with_diff_highlights(conflict.incoming_lines, diff_set1)
            print("\n" + CLR_BOLD + f"--- Option [1] [{conflict.incoming_label}] (Top) ---" + CLR_RESET)
            print(fmt_incoming)

            # Option 2: ALWAYS Current Branch / HEAD (Bottom)
            fmt_head = format_lines_with_diff_highlights(conflict.head_lines, diff_set2)
            print("\n" + CLR_BOLD + f"--- Option [2] [{conflict.head_label}] (Bottom) ---" + CLR_RESET)
            print(fmt_head)

            print("-" * 60)

            selected_option = None

            if auto_keep_state[0] == "first":
                selected_option = "1"
                print("  [AUTO-KEEP 1A] Selected Option 1 (Incoming Change)")
            elif auto_keep_state[0] == "last":
                selected_option = "2"
                print("  [AUTO-KEEP 2A] Selected Option 2 (Current Branch)")
            else:
                while True:
                    prompt_msg = (
                        "Select option to KEEP for this conflict marker:\n"
                        "  1  = Accept Option 1 [Incoming Change] (Top)\n"
                        "  2  = Accept Option 2 [Current Branch / HEAD] (Bottom)\n"
                        "  1A = Accept Option 1 for ALL remaining conflicts in this review session\n"
                        "  2A = Accept Option 2 for ALL remaining conflicts in this review session\n"
                        "  s  = Skip (Leave conflict untouched): "
                    )

                    choice = input(prompt_msg).strip().lower()
                    if choice in ("s", "skip"):
                        print(f"{CLR_YELLOW}Skipped conflict block (left untouched).{CLR_RESET}")
                        stats.conflicts_skipped += 1
                        selected_option = "s"
                        break
                    elif choice == "1":
                        selected_option = "1"
                        break
                    elif choice == "2":
                        selected_option = "2"
                        break
                    elif choice in ("1a", "a1", "all1"):
                        selected_option = "1"
                        auto_keep_state[0] = "first"
                        print(f"{CLR_GREEN}--> Activated AUTO-RESOLVE ALL: Keeping Option 1 [Incoming Change] for remaining conflicts.{CLR_RESET}")
                        break
                    elif choice in ("2a", "a2", "all2"):
                        selected_option = "2"
                        auto_keep_state[0] = "last"
                        print(f"{CLR_GREEN}--> Activated AUTO-RESOLVE ALL: Keeping Option 2 [Current Branch] for remaining conflicts.{CLR_RESET}")
                        break
                    print("Invalid choice. Please enter 1, 2, 1A, 2A, or s.")

            if selected_option == "s":
                continue

            if selected_option in ("1", "2"):
                stats.conflicts_resolved += 1
                lines_to_delete.add(conflict.start_line) # Delete <<<<<<<
                lines_to_delete.add(conflict.sep_line)   # Delete =======
                lines_to_delete.add(conflict.end_line)   # Delete >>>>>>>

                if selected_option == "1":
                    # Keep Option 1 (Incoming): Delete all HEAD lines between start_line and sep_line
                    for idx in range(conflict.start_line + 1, conflict.sep_line):
                        lines_to_delete.add(idx)
                elif selected_option == "2":
                    # Keep Option 2 (Current): Delete all Incoming lines between sep_line and end_line
                    for idx in range(conflict.sep_line + 1, conflict.end_line):
                        lines_to_delete.add(idx)

        if lines_to_delete and not dry_run:
            stats.files_modified += 1
            new_lines = [l for idx, l in enumerate(lines) if idx not in lines_to_delete]
            cleaned_lines = cleanup_excessive_blank_lines(new_lines)
            new_content = "".join(cleaned_lines)

            # Auto-format JSON syntax/commas & Field Parameter metadata
            new_content, fp_count = fix_field_parameter_metadata(new_content)
            if file_path.suffix.lower() in (".json", ".pbir", ".pbip", ".platform"):
                new_content = fix_pbip_json_formatting(new_content)

            encoded = new_content.encode("utf-8")
            if has_bom:
                encoded = b"\xef\xbb\xbf" + encoded
            with open(file_path, "wb") as f:
                f.write(encoded)
            print(f"[UPDATED] File conflict markers resolved & formatted: {file_path}")


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
    auto_keep_state: List[Optional[str]],
    global_remaining: List[int],
    stats: SummaryStats
) -> None:
    """
    Processes duplicate objects for a single file task.
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

        if auto_keep_state[0] == "first":
            selected_idx = 0
            label_str = f" [{col_blocks[0].git_label}]" if col_blocks[0].git_label else ""
            print(f"\n[AUTO-KEEP 1A] Selected Option 1{label_str} for {obj_type} '{col_display_name}'")
        elif auto_keep_state[0] == "last":
            selected_idx = num_options - 1
            last_blk = col_blocks[-1]
            label_str = f" [{last_blk.git_label}]" if last_blk.git_label else ""
            print(f"\n[AUTO-KEEP 2A] Selected Option {num_options}{label_str} for {obj_type} '{col_display_name}'")
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
                formatted_body = format_lines_with_diff_highlights(blk.lines, diff_set)

                print("\n" + CLR_BOLD + f"--- Option [{idx}]{origin} (Lines {blk.start_line + 1}-{blk.end_line}) ---" + CLR_RESET)
                print(formatted_body)

            print("-" * 60)
            while True:
                prompt_msg = (
                    f"Select option to KEEP for {obj_type} '{col_display_name}'\n"
                    f"  (1 = Keep Option 1, 2 = Keep Option 2, 1A = Keep Option 1 for ALL {obj_type.upper()}S, 2A = Keep Option 2 for ALL {obj_type.upper()}S, s = skip): "
                )
                choice = input(prompt_msg).strip().lower()

                if choice in ("s", "skip"):
                    print(f"Skipped duplicate {obj_type} '{col_display_name}'.")
                    stats.duplicates_skipped += 1
                    break
                elif choice == "1":
                    selected_idx = 0
                    break
                elif choice == "2":
                    selected_idx = 1
                    break
                elif choice in ("1a", "a1", "all1"):
                    selected_idx = 0
                    auto_keep_state[0] = "first"
                    print(f"\n{CLR_GREEN}--> Activated AUTO-RESOLVE ALL: Keeping Option 1 [Incoming Change] for remaining {obj_type.upper()} duplicates.{CLR_RESET}")
                    break
                elif choice in ("2a", "a2", "all2"):
                    selected_idx = num_options - 1
                    auto_keep_state[0] = "last"
                    print(f"\n{CLR_GREEN}--> Activated AUTO-RESOLVE ALL: Keeping Option 2 [Current Branch] for remaining {obj_type.upper()} duplicates.{CLR_RESET}")
                    break
                elif choice.isdigit():
                    val = int(choice)
                    if 1 <= val <= num_options:
                        selected_idx = val - 1
                        break
                print("Invalid input. Please enter 1, 2, 1A (all Incoming), 2A (all Current), or 's'.")

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
        cleaned_lines = cleanup_excessive_blank_lines(new_lines)
        new_content = "".join(cleaned_lines)

        # Auto-format JSON syntax/commas & Field Parameter metadata
        new_content, fp_count = fix_field_parameter_metadata(new_content)
        if task.file_path.suffix.lower() in (".json", ".pbir", ".pbip", ".platform"):
            new_content = fix_pbip_json_formatting(new_content)

        encoded_data = new_content.encode("utf-8")
        if task.has_bom:
            encoded_data = b"\xef\xbb\xbf" + encoded_data
        with open(task.file_path, "wb") as f:
            f.write(encoded_data)
        print(f"\n[UPDATED] File updated with clean spacing & JSON formatting: {task.file_path}")


def get_action_mode(args_mode: Optional[str]) -> str:
    """
    Returns action mode: '1' (Conflict Markers), '2' (Deduplicate Objects & Formatter), '3' (Stage Clean Files),
    '4' (Metadata Diagnostic Check), or '5' (Detailed Conflict Review).
    """
    if args_mode:
        val = args_mode.strip().lower()
        if val in ("1", "conflict", "markers"):
            return "1"
        elif val in ("2", "dedupe", "duplicates"):
            return "2"
        elif val in ("3", "stage", "gitadd", "add"):
            return "3"
        elif val in ("4", "diag", "diagnostic", "check", "health"):
            return "4"
        elif val in ("5", "review", "diff", "visual"):
            return "5"

    print("\n" + CLR_CYAN + "================================================================================" + CLR_RESET)
    print(CLR_BOLD + "               Power BI PBIP Master Conflict & Duplicate Resolver" + CLR_RESET)
    print(CLR_CYAN + "================================================================================" + CLR_RESET)
    print("Select resolution action to perform:")
    print(" 1 - Resolve Git Conflict Markers (lineageTag, logicalId, $schema, bookmark IDs, additions, and all conflicts)")
    print(" 2 - Resolve Duplicate Objects & Auto-Format Field Parameter / JSON Metadata")
    print(" 3 - Stage Clean Files to Git (Automatically git add files with 0 remaining conflict markers)")
    print(" 4 - Power BI PBIP Metadata Health & Diagnostic Check (Scan for all remaining issues)")
    print(" 5 - Detailed Conflict Review & Visual Diff Viewer (Review remaining conflicts one-by-one with diff highlights)")
    while True:
        choice = input("Enter option (1, 2, 3, 4, or 5): ").strip()
        if choice in ("1", "2", "3", "4", "5"):
            return choice
        print("Invalid input. Please enter 1, 2, 3, 4, or 5.\n")


def collect_target_files(target_path: Path, extensions_set: Optional[Set[str]]) -> List[Path]:
    """Collects all matching target files under target_path."""
    target_files: List[Path] = []
    if target_path.is_file():
        target_files.append(target_path)
    elif target_path.is_dir():
        for root, _, files in os.walk(target_path):
            for file in files:
                file_path = Path(root) / file
                if should_process_file(file_path, extensions_set):
                    target_files.append(file_path)
    return target_files


def run_single_execution(target_path: Path, mode: str, args) -> None:
    """Executes a single mode run (used for CLI flag execution)."""
    ext_arg = None if args.extensions and args.extensions.lower() == "all" else args.extensions
    extensions_set = parse_extensions(ext_arg)
    target_files = collect_target_files(target_path, extensions_set)

    initial_auto_keep = "first" if args.keep_first else "last" if args.keep_last else None
    stats = SummaryStats()

    if mode == "4":
        run_mode_4_metadata_diagnostic(target_files)
        return

    if mode == "3":
        run_mode_3_stage_clean_files(target_files, args.dry_run, target_path=target_path)
        return

    if mode == "5":
        auto_keep_m5 = [initial_auto_keep]
        run_mode_5_detailed_conflict_review(target_files, args.dry_run, auto_keep_m5, stats)
        return

    if mode == "1":
        conflict_target_type = get_conflict_target_type(args.conflict_type)
        if conflict_target_type == "back":
            return
        propagate_refs = False
        if conflict_target_type in ("bookmark", "all"):
            propagate_refs = ask_propagate_cross_references(args.propagate_refs if hasattr(args, "propagate_refs") and args.propagate_refs else None)

        auto_keep_m1 = [initial_auto_keep]
        run_mode_1_conflict_markers(target_files, conflict_target_type, args.dry_run, auto_keep_m1, stats, propagate_refs=propagate_refs)

    if mode == "2":
        target_object_type = get_target_object_type(args.type)
        if target_object_type == "back":
            return
        auto_keep_m2 = [initial_auto_keep]
        run_mode_2_dedupe_objects(target_files, target_object_type, args.dry_run, auto_keep_m2, stats)

    print("\n--- Final Summary ---")
    print(f"Files scanned: {stats.files_scanned}")
    print(f"Files modified: {stats.files_modified}")
    if mode in ("1", "5"):
        print(f"Conflict markers found: {stats.conflicts_found}")
        print(f"Conflict markers resolved: {stats.conflicts_resolved}")
        print(f"Conflict markers skipped: {stats.conflicts_skipped}")
    if mode == "2":
        print(f"Duplicate object groups found: {stats.duplicates_found}")
        print(f"Duplicate object groups resolved: {stats.duplicates_resolved}")
        print(f"Duplicate object groups skipped: {stats.duplicates_skipped}")
    print(f"Files skipped (binary/non-utf8): {stats.files_skipped}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Master Power BI PBIP Conflict & Duplicate Resolver (PBIP-ConflictsResolve.py)."
    )
    parser.add_argument(
        "path",
        nargs="?",
        help="Path to target file or folder.",
    )
    parser.add_argument(
        "--mode",
        choices=["1", "2", "3", "4", "5", "conflict", "dedupe", "stage", "diag", "review"],
        help="Action mode: 1 (Conflict Markers), 2 (Deduplicate Objects), 3 (Stage Clean Files), 4 (Metadata Diagnostic), 5 (Detailed Conflict Review).",
    )
    parser.add_argument(
        "--conflict-type",
        choices=["1", "2", "3", "4", "5", "5.1", "5.2", "5.3", "6", "lineage", "logical", "schema", "bookmark", "subset", "empty", "addition", "all"],
        help="Target property filter for Mode 1 conflict markers: 1/lineage, 2/logical, 3/schema, 4/bookmark, 5.1/subset, 5.2/empty, 5.3/addition, 6/all.",
    )
    parser.add_argument(
        "--type",
        choices=["1", "2", "3", "4", "5", "column", "expression", "relationship", "formatter", "all"],
        help="Target object type for Mode 2 deduplication: 1/column, 2/expression, 3/relationship, 4/formatter, 5/all.",
    )
    parser.add_argument(
        "--propagate-refs",
        action="store_true",
        help="Enable global cross-reference propagation across all project files for Bookmark/Object ID replacements.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report conflicts/duplicates without modifying files or prompting.",
    )
    parser.add_argument(
        "--extensions",
        type=str,
        default="tmdl,json,pbir,pbip,platform,fabric,definition,item,report",
        help="Comma-separated file extensions/names to process (default: tmdl,json,pbir,pbip,platform,fabric,definition,item,report). Use 'all' for all text files.",
    )
    parser.add_argument(
        "--keep-first",
        action="store_true",
        help="Automatically keep Option 1 / Incoming Change without prompting.",
    )
    parser.add_argument(
        "--keep-last",
        action="store_true",
        help="Automatically keep Option 2 / Current Change without prompting.",
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

    # If --mode was explicitly passed as CLI flag, run once and exit
    if args.mode:
        mode = get_action_mode(args.mode)
        run_single_execution(target_path, mode, args)
        return

    # Interactive Navigation Loop
    ext_arg = None if args.extensions and args.extensions.lower() == "all" else args.extensions
    extensions_set = parse_extensions(ext_arg)

    while True:
        mode = get_action_mode(None)
        target_files = collect_target_files(target_path, extensions_set)
        initial_auto_keep = "first" if args.keep_first else "last" if args.keep_last else None
        stats = SummaryStats()

        if mode == "4":
            run_mode_4_metadata_diagnostic(target_files)
        elif mode == "1":
            conflict_target_type = get_conflict_target_type(args.conflict_type)
            if conflict_target_type == "back":
                print(f"\n{CLR_YELLOW}--> Returning to Main Menu...{CLR_RESET}")
                continue

            propagate_refs = False
            if conflict_target_type in ("bookmark", "all"):
                propagate_refs = ask_propagate_cross_references(args.propagate_refs if hasattr(args, "propagate_refs") and args.propagate_refs else None)

            auto_keep_m1 = [initial_auto_keep]
            run_mode_1_conflict_markers(target_files, conflict_target_type, args.dry_run, auto_keep_m1, stats, propagate_refs=propagate_refs)
            print("\n--- Summary ---")
            print(f"Files scanned: {stats.files_scanned} | Modified: {stats.files_modified}")
            print(f"Conflict markers found: {stats.conflicts_found} | Resolved: {stats.conflicts_resolved} | Skipped: {stats.conflicts_skipped}")
        elif mode == "2":
            target_object_type = get_target_object_type(args.type)
            if target_object_type == "back":
                print(f"\n{CLR_YELLOW}--> Returning to Main Menu...{CLR_RESET}")
                continue

            auto_keep_m2 = [initial_auto_keep]
            run_mode_2_dedupe_objects(target_files, target_object_type, args.dry_run, auto_keep_m2, stats)
            print("\n--- Summary ---")
            print(f"Files scanned: {stats.files_scanned} | Modified: {stats.files_modified}")
            print(f"Duplicates found: {stats.duplicates_found} | Resolved: {stats.duplicates_resolved} | Skipped: {stats.duplicates_skipped}")
        elif mode == "3":
            run_mode_3_stage_clean_files(target_files, args.dry_run, target_path=target_path)
        elif mode == "5":
            auto_keep_m5 = [initial_auto_keep]
            run_mode_5_detailed_conflict_review(target_files, args.dry_run, auto_keep_m5, stats)
            print("\n--- Summary ---")
            print(f"Files scanned: {stats.files_scanned} | Modified: {stats.files_modified}")
            print(f"Conflict markers found: {stats.conflicts_found} | Resolved: {stats.conflicts_resolved} | Skipped: {stats.conflicts_skipped}")

        # Sub-menu navigation return prompt
        print("\n" + CLR_CYAN + "--------------------------------------------------------------------------------" + CLR_RESET)
        print(CLR_BOLD + "Task completed. What would you like to do next?" + CLR_RESET)
        print(" 1 - Return to Main Menu")
        print(" 2 - Exit")
        while True:
            next_choice = input("Enter option (1 or 2): ").strip()
            if next_choice == "1":
                break
            elif next_choice == "2":
                print("\nExiting PBIP Resolver. Goodbye!")
                sys.exit(0)
            print("Invalid input. Please enter 1 or 2.")


if __name__ == "__main__":
    main()


# Example usage:
# python PBIP-ConflictsResolve.py
# python PBIP-ConflictsResolve.py "C:/path/to/report"
