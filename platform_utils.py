"""Small cross-platform helpers for launching files and desktop apps."""

from __future__ import annotations

import os
from pathlib import Path
import shutil
import subprocess
import sys


def open_path(path: str | os.PathLike[str]) -> None:
    """Open a file or folder with the operating system's default application."""
    target = str(Path(path).expanduser().resolve())
    if sys.platform == "win32":
        os.startfile(target)  # type: ignore[attr-defined]
        return
    if sys.platform == "darwin":
        subprocess.Popen(["open", target])
        return

    opener = shutil.which("xdg-open")
    if opener is None:
        raise RuntimeError("未找到可用的文件打开程序（xdg-open）")
    subprocess.Popen([opener, target])


def launch_word() -> None:
    """Launch Microsoft Word, or WPS Office as the macOS fallback."""
    if sys.platform == "win32":
        subprocess.Popen(["cmd", "/c", "start", "", "winword"])
        return
    if sys.platform == "darwin":
        errors: list[str] = []
        app_targets = (
            ["open", "-a", "Microsoft Word"],
            ["open", "-b", "com.kingsoft.wpsoffice.mac"],
            ["open", "/Applications/wpsoffice.app"],
        )
        for command in app_targets:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
            )
            if result.returncode == 0:
                return
            detail = (result.stderr or result.stdout).strip()
            if detail:
                errors.append(detail)
        raise RuntimeError(errors[-1] if errors else "未找到 Microsoft Word 或 WPS Office")

    word = shutil.which("libreoffice") or shutil.which("soffice")
    if word is None:
        raise RuntimeError("未找到 Microsoft Word 或 LibreOffice")
    subprocess.Popen([word, "--writer"])
