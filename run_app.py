"""Run the Streamlit application with the project's local-only default."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    project_dir = Path(__file__).resolve().parent / "06-Haziran"
    launcher = project_dir / "run.py"
    if not launcher.is_file():
        raise FileNotFoundError(f"Project launcher was not found: {launcher}")

    return subprocess.run(
        [sys.executable, str(launcher), *sys.argv[1:]],
        cwd=project_dir,
        check=False,
    ).returncode


if __name__ == "__main__":
    raise SystemExit(main())
