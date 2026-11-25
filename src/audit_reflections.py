#!/usr/bin/env python3
"""
Step 2: Shallow-clone a GitHub reflections repository and audit unique classes (1..12)
per student ID, filtering out 1-byte placeholders and deduplicating .md/.docx pairs.
Author: Gautam Gupta (https://www.linkedin.com/in/higautam/)
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

ORDINALS = {
    "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5, "sixth": 6,
    "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10, "eleventh": 11, "twelfth": 12,
}
IGNORE_FILES = ("ai-use-disclosure", "readme", "read me", ".ds_store", ".gitkeep", "project.md", "test.md")


def read_text(path: str) -> str:
    if path.lower().endswith(".docx"):
        try:
            with zipfile.ZipFile(path) as z:
                tree = ET.fromstring(z.read("word/document.xml"))
            return " ".join(t.text for t in tree.iter() if t.tag.endswith("}t") and t.text)
        except Exception:
            return ""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        return ""


def extract_classes(fname: str, text: str, max_classes: int = 12) -> set[int]:
    base = os.path.basename(fname).lower()
    if any(k in base for k in IGNORE_FILES):
        return set()

    nums: set[int] = set()
    # Range pattern: Session1_to_Session7
    rm = re.search(r"(?:session|class)s?[ _-]*0?(\d{1,2})[ _-]*to[ _-]*(?:session|class)?[ _-]*0?(\d{1,2})", base)
    if rm:
        for n in range(int(rm.group(1)), int(rm.group(2)) + 1):
            if 1 <= n <= max_classes:
                nums.add(n)
        return nums

    for word, cnum in ORDINALS.items():
        if word in base and cnum <= max_classes:
            nums.add(cnum)

    for m in re.finditer(r"(?:class|session|reflection|reflections|day|week|lec)[ _-]*(?<!\d)0?(\d{1,2})(?!\d)", base):
        n = int(m.group(1))
        if 1 <= n <= max_classes:
            nums.add(n)

    if not nums and text:
        for m in re.finditer(r"(?:class|session|lecture|day|week)[ _:-]*(?<!\d)0?(\d{1,2})(?!\d)", text.lower()):
            n = int(m.group(1))
            if 1 <= n <= max_classes:
                nums.add(n)
    return nums


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit class reflections in a GitHub repo.")
    parser.add_argument("--repo", required=True, help="owner/repo slug (e.g., ORG/Class-Reflections)")
    parser.add_argument("--id-pattern", default=r"(202\d[a-z]{2}\d+)", help="Regex to extract student ID from path")
    parser.add_argument("--max-classes", type=int, default=12, help="Total expected classes")
    parser.add_argument("--out", default="reflections_audit.json", help="Output JSON path")
    args = parser.parse_args()

    token = os.environ.get("GH_TOKEN")
    if not token:
        print("Error: GH_TOKEN environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    tmp_dir = tempfile.mkdtemp(prefix="refl_audit_")
    try:
        repo_url = f"https://{token}@github.com/{args.repo}.git"
        subprocess.run(["git", "clone", "--depth", "1", repo_url, tmp_dir], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        by_sid = defaultdict(list)
        id_re = re.compile(args.id_pattern, re.IGNORECASE)
        for dirpath, _, filenames in os.walk(tmp_dir):
            if "/.git" in dirpath or dirpath.endswith("/.git"):
                continue
            for fn in filenames:
                if fn in (".DS_Store", ".gitkeep"):
                    continue
                full = os.path.join(dirpath, fn)
                rel = os.path.relpath(full, tmp_dir)
                m = id_re.search(rel)
                if m:
                    by_sid[m.group(1).upper()].append((rel, os.path.getsize(full), full))

        audit = {}
        for sid, files in sorted(by_sid.items()):
            classes: set[int] = set()
            for rel, size, full in files:
                if size <= 15:
                    continue
                is_pdf = full.lower().endswith(".pdf")
                text = "" if is_pdf else read_text(full)
                if not is_pdf and len(text.split()) < 15:
                    continue
                classes.update(extract_classes(rel, text, args.max_classes))

            cnt = len(classes)
            audit[sid] = {
                "classes": sorted(classes),
                "count": cnt,
                "marks": round((cnt / args.max_classes) * 10.0, 2),
                "has_any_file": len(files) > 0,
            }

        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(audit, f, indent=2)
        print(f"Audited {len(audit)} students -> {args.out}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
