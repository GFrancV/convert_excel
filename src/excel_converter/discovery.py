from pathlib import Path

OLD_FORMATS = {".xls"}


def find_files(input_dir: Path, recursive: bool) -> list:
    """Return all .xls files in input_dir."""
    glob = input_dir.rglob if recursive else input_dir.glob
    return [f for f in glob("*") if f.suffix.lower() in OLD_FORMATS]


def build_tasks(files: list, input_dir: Path, output_dir: Path) -> list:
    """Map each source file to its destination path."""
    return [
        (src, output_dir / src.relative_to(input_dir).with_suffix(".xlsx"))
        for src in files
    ]
