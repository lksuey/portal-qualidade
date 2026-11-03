#!/usr/bin/env python3
"""Pre-publication scan for sensitive content. Python standard library only.

Prints ONLY counts and file:line locations, never the matched values.
Exit code 0 = zero findings, 1 = findings, 2 = usage error.

    python scan_sensitive.py                       # scan this repository
    python scan_sensitive.py --names-file terms.txt  # also look for private terms

--names-file points to a text file OUTSIDE the repository (one term per line, '#'
comments allowed) with the proper names, internal system names and ids you do not
want published. Keep that file out of version control: it is the one place that
holds the sensitive strings. Matching is case and accent insensitive.

Categories: cpf, cnpj, email, phone, url, path, secret, google_id, private_term,
unexpected_file. This file is excluded from its own scan.
"""
import argparse
import os
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

SKIP_DIRS = {".git", "node_modules", "__pycache__", ".venv"}
TEXT_EXT = {".html", ".css", ".js", ".gs", ".py", ".md", ".json", ".txt", ".yml", ".yaml", ".gitignore", ".mjs", ".cjs", ".toml", ".cfg"}
SAFE_EMAIL_DOMAINS = ("example.com", "example.org", "example.net", "example.invalid", "users.noreply.github.com")

CPF_FMT = re.compile(r"(?<!\d)\d{3}\.\d{3}\.\d{3}-\d{2}(?!\d)")
CNPJ_FMT = re.compile(r"(?<!\d)\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}(?!\d)")
DIGITS11 = re.compile(r"(?<![\dA-Za-z+/_-])\d{11}(?![\dA-Za-z+/_-])")
DIGITS14 = re.compile(r"(?<![\dA-Za-z+/_-])\d{14}(?![\dA-Za-z+/_-])")
EMAIL = re.compile(r"[A-Za-z0-9._%+-]+@([A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,})")
PHONE = re.compile(r"(?<![\dA-Za-z+/_-])(?:\+?55[\s.-]?)?\(?[1-9][1-9]\)?[\s.-]?9?\d{4}[\s.-]?\d{4}(?![\dA-Za-z+/_-])")
URL_PATTERNS = [
    re.compile(r"script\.google\.com/macros", re.I),
    re.compile(r"docs\.google\.com/spreadsheets", re.I),
    re.compile(r"docs\.google\.com/document", re.I),
    re.compile(r"drive\.google\.com/(file|drive)", re.I),
    re.compile(r"https?://[^\s\"'<>)]*\.(?:local|internal|corp|lan)(?:[:/]|\b)", re.I),
]
PATH_PATTERNS = [
    re.compile(r"[A-Za-z]:\\\\?[A-Za-z0-9_ .$-]+\\\\?[A-Za-z0-9_ .$-]"),   # drive letter paths
    re.compile(r"[A-Za-z]:/(?:Users|Documents and Settings|Windows|Program Files)", re.I),
    re.compile(r"\\\\[A-Za-z0-9_.-]+\\[A-Za-z0-9_$.-]+"),                 # UNC
    re.compile(r"(?<![\w.])/(?:Users|home)/[A-Za-z0-9._-]+/"),
]
SECRET_PATTERNS = [
    re.compile(r"AKIA[0-9A-Z]{16}"),
    re.compile(r"AIza[0-9A-Za-z_-]{35}"),
    re.compile(r"ya29\.[0-9A-Za-z_-]{20,}"),
    re.compile(r"-----BEGIN [A-Z ]*PRIVATE KEY-----"),
    re.compile(r"\"private_key\"\s*:"),
    re.compile(r"eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{5,}"),
    re.compile(r"(?i)(password|passwd|senha|secret|token|api[_-]?key|apikey)['\"]?\s*[:=]\s*['\"][^'\"\s]{6,}['\"]"),
    re.compile(r"(?i)\b(AG_EMAIL|AG_PASSWORD)\b\s*="),
]
GOOGLE_ID = re.compile(r"(?<![A-Za-z0-9_-])[A-Za-z0-9_-]{40,}(?![A-Za-z0-9_-])")


def cpf_valid(d):
    if len(set(d)) == 1:
        return False
    for n in (9, 10):
        s = sum(int(d[i]) * (n + 1 - i) for i in range(n))
        if (s * 10 % 11) % 10 != int(d[n]):
            return False
    return True


def cnpj_valid(d):
    if len(set(d)) == 1:
        return False
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    w2 = [6] + w1
    for ws, pos in ((w1, 12), (w2, 13)):
        s = sum(int(d[i]) * ws[i] for i in range(pos))
        r = s % 11
        if int(d[pos]) != (0 if r < 2 else 11 - r):
            return False
    return True


def placeholder_assignment(matched):
    """name = 'value' matches are ignored when the value looks like a plain word
    (lowercase letters and dashes only), e.g. a placeholder or a storage key name."""
    m = re.search(r"[:=]\s*['\"]([^'\"\s]+)['\"]\s*$", matched)
    if not m:
        return False
    return re.fullmatch(r"[a-z]+(?:[-_][a-z]+)*", m.group(1)) is not None


def fold(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)).lower()


def load_terms(path):
    terms = []
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            terms.append(fold(line))
    return terms


def iter_files(root, self_path):
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for name in filenames:
            p = Path(dirpath) / name
            if p.resolve() == self_path:
                continue
            yield p


def scan(root, terms):
    root = Path(root).resolve()
    findings = defaultdict(list)  # category -> [(relpath, line)]
    for path in iter_files(root, Path(__file__).resolve()):
        rel = path.relative_to(root).as_posix()
        ext = path.suffix.lower() if path.suffix else path.name.lower()
        if ext not in TEXT_EXT:
            findings["unexpected_file"].append((rel, 0))
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            findings["unexpected_file"].append((rel, 0))
            continue
        for lineno, line in enumerate(text.split("\n"), 1):
            is_b64 = "base64," in line
            for m in CPF_FMT.finditer(line):
                findings["cpf"].append((rel, lineno))
            for m in DIGITS11.finditer(line):
                if cpf_valid(m.group()):
                    findings["cpf"].append((rel, lineno))
            for m in CNPJ_FMT.finditer(line):
                findings["cnpj"].append((rel, lineno))
            for m in DIGITS14.finditer(line):
                if cnpj_valid(m.group()):
                    findings["cnpj"].append((rel, lineno))
            for m in EMAIL.finditer(line):
                if not m.group(1).lower().endswith(SAFE_EMAIL_DOMAINS):
                    findings["email"].append((rel, lineno))
            if not is_b64:
                for m in PHONE.finditer(line):
                    findings["phone"].append((rel, lineno))
            for pat in URL_PATTERNS:
                if pat.search(line):
                    findings["url"].append((rel, lineno))
            for pat in PATH_PATTERNS:
                if pat.search(line):
                    findings["path"].append((rel, lineno))
            for pat in SECRET_PATTERNS:
                m = pat.search(line)
                if m and not placeholder_assignment(m.group()):
                    findings["secret"].append((rel, lineno))
            if not is_b64:
                for m in GOOGLE_ID.finditer(line):
                    tok = m.group()
                    if re.search(r"\d", tok) and re.search(r"[A-Za-z]", tok):
                        findings["google_id"].append((rel, lineno))
            if terms:
                folded = fold(line)
                if any(t in folded for t in terms):
                    findings["private_term"].append((rel, lineno))
    return findings


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("root", nargs="?", default=str(Path(__file__).resolve().parent))
    ap.add_argument("--names-file", help="private terms file kept outside the repository")
    args = ap.parse_args()
    terms = load_terms(args.names_file) if args.names_file else []
    if args.names_file and not terms:
        print("names file is empty", file=sys.stderr)
        return 2
    findings = scan(args.root, terms)
    categories = ["cpf", "cnpj", "email", "phone", "url", "path", "secret", "google_id", "private_term", "unexpected_file"]
    total = 0
    print("Scan of %s (%d private terms loaded)" % (Path(args.root).resolve().name, len(terms)))
    for cat in categories:
        hits = sorted(set(findings.get(cat, [])))
        total += len(hits)
        print("  %-16s %d" % (cat, len(hits)))
        for rel, line in hits[:50]:
            print("      %s:%s" % (rel, line))
    print("TOTAL findings: %d" % total)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
