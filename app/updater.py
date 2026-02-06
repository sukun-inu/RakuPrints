from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def _copy_item(src: Path, dest: Path) -> None:
    if src.is_dir():
        shutil.copytree(src, dest, dirs_exist_ok=True)
    else:
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)


def apply_update(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--target", required=True)
    parser.add_argument("--source", default="")
    parser.add_argument("--restart", action="store_true")
    args = parser.parse_args(argv)

    target_dir = Path(args.target).resolve()
    source_dir = Path(args.source).resolve() if args.source else Path(sys.executable).parent

    if not source_dir.exists() or not target_dir.exists():
        return 1

    for item in source_dir.iterdir():
        if item.name in {"config", "logging"}:
            continue
        _copy_item(item, target_dir / item.name)

    if args.restart:
        exe_name = Path(sys.executable).name
        target_exe = target_dir / exe_name
        if target_exe.exists():
            try:
                subprocess.Popen([str(target_exe)])
            except Exception:
                pass
    return 0
