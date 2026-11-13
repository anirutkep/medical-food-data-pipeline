# -*- coding: utf-8 -*-
"""
Workspace Hygiene & Cache Cleanup Engine.

Removes transient cache directories (__pycache__, .ocr_cache) and OS-generated metadata
files (.DS_Store, Thumbs.db) to ensure clean repository packaging and deliverables hygiene.
"""

from __future__ import annotations
import argparse
import os
import shutil
import sys
from pathlib import Path
from typing import List, Optional


def cleanup_workspace(root: str | Path, dry_run: bool = False) -> List[Path]:
    """
    Remove temporary cache directories and OS junk files.

    Target directories: __pycache__, .ocr_cache
    Target files: .DS_Store, Thumbs.db

    Returns sorted list of Path objects cleaned (or detected in dry-run).
    """
    root_path = Path(root).resolve()
    targets: List[Path] = []

    target_dir_names = {"__pycache__", ".ocr_cache"}
    target_file_names = {".DS_Store", "Thumbs.db"}

    for dirpath, dirnames, filenames in os.walk(root_path, topdown=True):
        # 1. Check files
        for fname in filenames:
            if fname in target_file_names:
                fpath = Path(dirpath) / fname
                targets.append(fpath)
                if not dry_run:
                    try:
                        fpath.unlink()
                    except OSError:
                        pass

        # 2. Check directories to remove & prune from dirnames
        dirs_to_prune = []
        for dname in list(dirnames):
            if dname in target_dir_names:
                dpath = Path(dirpath) / dname
                targets.append(dpath)
                dirs_to_prune.append(dname)
                if not dry_run:
                    shutil.rmtree(dpath, ignore_errors=True)

        for dname in dirs_to_prune:
            dirnames.remove(dname)

    return sorted(targets)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Workspace Cleanup & Hygiene Tool")
    parser.add_argument("root", nargs="?", default=".", help="Root directory to clean (default: current directory)")
    parser.add_argument("--dry-run", action="store_true", help="Report paths without deleting")
    args = parser.parse_args(argv)

    targets = cleanup_workspace(args.root, dry_run=args.dry_run)
    mode_str = "[DRY-RUN] Would remove" if args.dry_run else "Removed"
    print(f"=== Workspace Cleanup: {mode_str} {len(targets)} item(s) ===")
    for t in targets:
        print(f"  {t}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
