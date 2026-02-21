#!/usr/bin/env python3
"""Install/update the citation-resolver skill into Codex home."""

from __future__ import annotations

import argparse
import json
import os
import shutil
from pathlib import Path


def _default_codex_home() -> Path:
    env_home = os.environ.get("CODEX_HOME")
    if env_home:
        return Path(env_home).expanduser().resolve()
    return (Path.home() / ".codex").resolve()


def install_skill(codex_home: Path, skill_name: str, dry_run: bool) -> dict:
    source_root = Path(__file__).resolve().parent
    source_script = source_root / "docx_zotero_integrator.py"
    source_skill = source_root / "SKILL.md"
    if not source_script.exists() or not source_skill.exists():
        raise FileNotFoundError("Missing docx_zotero_integrator.py or SKILL.md in repository root")

    install_root = codex_home / "skills" / skill_name
    targets = {
        "script": install_root / "docx_zotero_integrator.py",
        "skill": install_root / "SKILL.md",
    }

    if not dry_run:
        install_root.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_script, targets["script"])
        shutil.copy2(source_skill, targets["skill"])

    return {
        "codex_home": str(codex_home),
        "skill_name": skill_name,
        "install_root": str(install_root),
        "dry_run": dry_run,
        "files": {
            "docx_zotero_integrator.py": str(targets["script"]),
            "SKILL.md": str(targets["skill"]),
        },
    }


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Install/update citation-resolver skill into Codex")
    p.add_argument(
        "--codex-home",
        default=str(_default_codex_home()),
        help="Codex home folder (default: CODEX_HOME or ~/.codex)",
    )
    p.add_argument(
        "--skill-name",
        default="citation-resolver",
        help="Skill folder name under <codex_home>/skills (default: citation-resolver)",
    )
    p.add_argument("--dry-run", action="store_true", help="Show install plan without writing files")
    return p


def main() -> int:
    args = build_parser().parse_args()
    report = install_skill(
        codex_home=Path(args.codex_home).expanduser().resolve(),
        skill_name=args.skill_name,
        dry_run=args.dry_run,
    )
    print(json.dumps(report, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
