#!/usr/bin/env python3
"""
Git Merge Conflict Resolver for Power BI PBIP lineageTag / sourceLineageTag properties.
"""

import argparse
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Set, Tuple


# Regex pattern to match lineageTag or sourceLineageTag key-value property lines.
# Supports unquoted, double-quoted, or single-quoted keys with case-insensitivity.
# Examples:
#   lineageTag: "abc"
#   "lineageTag": "abc",
#   sourceLineageTag: xyz
#   'sourceLineageTag': "123",
LINEAGE_TAG_PATTERN = re.compile(
    r'^\s*(?:"(?:lineageTag|sourceLineageTag)"|\'(?:lineageTag|sourceLineageTag)\'|(?:lineageTag|sourceLineageTag)\b)\s*:',
    re.IGNORECASE,
)


@dataclass
class SummaryStats:
    files_scanned: int = 0
    files_modified: int = 0
    conflicts_resolved: int = 0
    conflicts_remaining: int = 0
    files_skipped: int = 0


def is_lineage_tag_line(line: str) -> bool:
    """Check if a line is a key-value property line for lineageTag or sourceLineageTag."""
    return bool(LINEAGE_TAG_PATTERN.match(line))


def resolve_conflicts_in_content(
    content: str, mode: int
) -> Tuple[str, int, int]:
    """
    Parses content for Git conflict blocks and resolves eligible blocks.

    Args:
        content: File text content.
        mode: 1 for incoming change, 2 for current change.

    Returns:
        Tuple of (new_content, conflicts_resolved, conflicts_remaining).
    """
    lines = content.splitlines(keepends=True)
    output_lines = []

    # States
    OUTSIDE = 0
    IN_CURRENT = 1
    IN_INCOMING = 2

    state = OUTSIDE
    current_side: list[str] = []
    incoming_side: list[str] = []
    start_marker = ""
    separator_marker = ""

    resolved_count = 0
    remaining_count = 0

    for line in lines:
        stripped = line.strip()
        if state == OUTSIDE:
            if stripped.startswith("<<<<<<<"):
                start_marker = line
                current_side = []
                incoming_side = []
                state = IN_CURRENT
            else:
                output_lines.append(line)

        elif state == IN_CURRENT:
            if stripped.startswith("======="):
                separator_marker = line
                state = IN_INCOMING
            elif stripped.startswith("<<<<<<<"):
                # Unclosed conflict marker encountered before starting new one
                remaining_count += 1
                output_lines.append(start_marker)
                output_lines.extend(current_side)
                start_marker = line
                current_side = []
            else:
                current_side.append(line)

        elif state == IN_INCOMING:
            if stripped.startswith(">>>>>>>"):
                end_marker = line
                state = OUTSIDE

                # Check eligibility
                curr_non_empty = [l for l in current_side if l.strip()]
                inc_non_empty = [l for l in incoming_side if l.strip()]

                is_eligible = (
                    len(curr_non_empty) == 1
                    and len(inc_non_empty) == 1
                    and is_lineage_tag_line(curr_non_empty[0])
                    and is_lineage_tag_line(inc_non_empty[0])
                )

                if is_eligible:
                    resolved_count += 1
                    if mode == 1:  # incoming
                        output_lines.extend(incoming_side)
                    else:  # current (2)
                        output_lines.extend(current_side)
                else:
                    remaining_count += 1
                    output_lines.append(start_marker)
                    output_lines.extend(current_side)
                    output_lines.append(separator_marker)
                    output_lines.extend(incoming_side)
                    output_lines.append(end_marker)

            elif stripped.startswith("<<<<<<<"):
                # Malformed block: unexpected start marker while inside incoming
                remaining_count += 1
                output_lines.append(start_marker)
                output_lines.extend(current_side)
                output_lines.append(separator_marker)
                output_lines.extend(incoming_side)
                start_marker = line
                current_side = []
                incoming_side = []
                state = IN_CURRENT
            else:
                incoming_side.append(line)

    if state != OUTSIDE:
        # Handles unclosed conflict block at EOF
        remaining_count += 1
        output_lines.append(start_marker)
        output_lines.extend(current_side)
        if state == IN_INCOMING:
            output_lines.append(separator_marker)
            output_lines.extend(incoming_side)

    new_content = "".join(output_lines)
    return new_content, resolved_count, remaining_count


def process_file(
    file_path: Path, mode: int, dry_run: bool, stats: SummaryStats
) -> None:
    """Reads, processes, and writes a single file if modified."""
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

    new_content, resolved, remaining = resolve_conflicts_in_content(content, mode)
    stats.conflicts_resolved += resolved
    stats.conflicts_remaining += remaining

    if new_content != content:
        stats.files_modified += 1
        if not dry_run:
            encoded_data = new_content.encode("utf-8")
            if has_bom:
                encoded_data = b"\xef\xbb\xbf" + encoded_data
            with open(file_path, "wb") as f:
                f.write(encoded_data)
        action = "[DRY-RUN] Would modify" if dry_run else "Modified"
        print(f"{action}: {file_path} (Resolved: {resolved}, Remaining: {remaining})")


def get_resolution_mode() -> int:
    """Prompts user to select resolution mode until valid input is given."""
    while True:
        choice = input(
            "Select resolution mode:\n 1 - Incoming change\n 2 - Current change\nEnter option (1 or 2): "
        ).strip()
        if choice in ("1", "2"):
            return int(choice)
        print("Invalid input. Please enter 1 or 2.\n")


def parse_extensions(ext_str: Optional[str]) -> Optional[Set[str]]:
    """Parses comma-separated extension string into a normalized set."""
    if not ext_str:
        return None
    exts = {e.strip().lstrip(".").lower() for e in ext_str.split(",") if e.strip()}
    return exts if exts else None


def should_process_file(file_path: Path, extensions: Optional[Set[str]]) -> bool:
    """Determines if a file matches the target extensions filter."""
    if extensions is None:
        return True
    ext = file_path.suffix.lstrip(".").lower()
    return ext in extensions


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Resolve Git merge conflicts in Power BI PBIP lineageTag/sourceLineageTag properties."
    )
    parser.add_argument(
        "path",
        nargs="?",
        help="Path to a target file or folder.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report changes without writing to files.",
    )
    parser.add_argument(
        "--extensions",
        type=str,
        help="Comma-separated file extensions to process (e.g. json,tmdl,pbir,pbip).",
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

    mode = get_resolution_mode()
    extensions_set = parse_extensions(args.extensions)

    stats = SummaryStats()

    if target_path.is_file():
        if should_process_file(target_path, extensions_set):
            process_file(target_path, mode, args.dry_run, stats)
        else:
            print(f"Skipped (extension mismatch): {target_path}")
    elif target_path.is_dir():
        for root, _, files in os.walk(target_path):
            for file in files:
                file_path = Path(root) / file
                if should_process_file(file_path, extensions_set):
                    process_file(file_path, mode, args.dry_run, stats)
    else:
        print(f"Error: Path '{target_path}' is neither a file nor a directory.", file=sys.stderr)
        sys.exit(1)

    print("\n--- Summary ---")
    print(f"Files scanned: {stats.files_scanned}")
    print(f"Files modified: {stats.files_modified}")
    print(f"Conflicts resolved: {stats.conflicts_resolved}")
    print(f"Unresolved conflict blocks remaining: {stats.conflicts_remaining}")
    print(f"Files skipped: {stats.files_skipped}")


if __name__ == "__main__":
    main()


# Example usage:
# python task.py
# python task.py "C:/path/to/folder" --dry-run
# python task.py "C:/path/to/folder" --extensions json,tmdl,pbir,pbip
