#!/usr/bin/env python3
"""
Bump the app version in every place that tracks it, so a release stays
in sync between the Python app, the Windows installer, and the git tag.

Usage:
    python build/release.py 1.8.0

Updates:
    scandocs_tool.py   APP_VERSION = "X.Y.Z"
    build/windows/installer.iss   #define AppVersion "X.Y.Z"

After running this script:
    1. Build the Windows installer:
         build\\windows\\build_windows.bat
       Output: build\\windows\\Output\\SpeedyScandocsSetup.exe

    2. Build the Mac app (on a Mac):
         pyinstaller build/mac/scandocs_mac.spec --clean
         create-dmg SpeedyScandocs.dmg dist/SpeedyScandocs.app

    3. Commit and tag:
         git commit -am "Release vX.Y.Z"
         git tag vX.Y.Z
         git push && git push --tags

    4. Publish the GitHub Release (auto-update reads this):
         gh release create vX.Y.Z \\
             build/windows/Output/SpeedyScandocsSetup.exe \\
             SpeedyScandocs.dmg \\
             --title "vX.Y.Z" --notes "What changed in this release"

Within 24 hours of that release (or on the next manual "Check for Updates"),
every installed copy of Speedy Scandocs will offer the update.
"""
from __future__ import annotations
import argparse
import pathlib
import re
import sys

REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent
APP_FILE = REPO_ROOT / "scandocs_tool.py"
ISS_FILE = REPO_ROOT / "build" / "windows" / "installer.iss"

VERSION_RE = re.compile(r"^\d+\.\d+\.\d+$")


def bump_app_py(version: str) -> None:
    text = APP_FILE.read_text(encoding="utf-8")
    new_text, n = re.subn(
        r'^APP_VERSION = "[^"]*"',
        f'APP_VERSION = "{version}"',
        text,
        count=1,
        flags=re.MULTILINE,
    )
    if n != 1:
        sys.exit(f"Could not find APP_VERSION line in {APP_FILE}")
    APP_FILE.write_text(new_text, encoding="utf-8")
    print(f"  Updated {APP_FILE.name}  APP_VERSION -> {version}")


def bump_iss(version: str) -> None:
    text = ISS_FILE.read_text(encoding="utf-8")
    new_text, n = re.subn(
        r'#define AppVersion\s+"[^"]*"',
        f'#define AppVersion   "{version}"',
        text,
        count=1,
    )
    if n != 1:
        sys.exit(f"Could not find #define AppVersion in {ISS_FILE}")
    ISS_FILE.write_text(new_text, encoding="utf-8")
    print(f"  Updated {ISS_FILE.name}  AppVersion -> {version}")


def main() -> None:
    p = argparse.ArgumentParser(description="Bump Speedy Scandocs version.")
    p.add_argument("version", help='New version, e.g. "1.8.0" (no "v" prefix).')
    args = p.parse_args()

    v = args.version.lstrip("vV").strip()
    if not VERSION_RE.match(v):
        sys.exit(f'Version must be MAJOR.MINOR.PATCH (got "{args.version}").')

    bump_app_py(v)
    bump_iss(v)

    print()
    print(f"Version bumped to {v}. Next steps:")
    print(f"  1. Build Windows:   build\\windows\\build_windows.bat")
    print(f"  2. Build Mac:       pyinstaller build/mac/scandocs_mac.spec --clean")
    print(f"  3. Tag + push:      git commit -am 'Release v{v}' && git tag v{v} && git push --tags")
    print(f"  4. GitHub release:  gh release create v{v} <installer.exe> <app.dmg> --title 'v{v}'")


if __name__ == "__main__":
    main()
