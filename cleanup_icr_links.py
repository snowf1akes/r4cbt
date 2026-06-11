#!/usr/bin/env python3
"""
cleanup_icr_links.py

Fixes the stale-`post` generator bug in ICR annotation sheets: every
thread-level OP block's hyperlink (column E) points to the LAST post in the
source JSON instead of its own post.

Logic per sheet:
  1. Walk rows, tracking post boundaries via "Post N" headers in column A.
  2. The post-header block's link (title row) is authoritative — the generator
     built it from op_info_by_post, which was always correct.
  3. Cross-check it against comment permalinks in the same post (their post ID
     must match). If the header link is missing, derive the post link from a
     comment permalink by stripping the comment-ID segment.
  4. Inside thread blocks ("Post N: Thread M"), any post-level link whose post
     ID differs from the authoritative one is rewritten (hyperlink + display
     value). Annotations and formatting are untouched.

Usage:
  python cleanup_icr_links.py sheet.xlsx [more.xlsx ...]
  python cleanup_icr_links.py --dir folder_of_sheets/
Output: <name>_fixed.xlsx next to each input (originals never modified).
"""
import argparse
import os
import re
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font

POST_HDR = re.compile(r"^Post (\d+)$")
THREAD_HDR = re.compile(r"^Post (\d+): Thread (\d+)$")
# group(1)=base post URL, group(2)=post ID, group(3)=comment ID (None if post-level)
REDDIT_URL = re.compile(
    r"^(https?://(?:www\.)?reddit\.com/r/[^/]+/comments/([a-z0-9]+)(?:/[^/]+)?/?)"
    r"(?:([a-z0-9]+)/?)?$"
)


def parse_link(url):
    """Return (post_id, comment_id, base_post_url) or (None, None, None)."""
    if not url:
        return None, None, None
    m = REDDIT_URL.match(url.strip())
    if not m:
        return None, None, None
    base = m.group(1)
    if not base.endswith("/"):
        base += "/"
    return m.group(2), m.group(3), base


def scan_posts(ws):
    """Map each row to its post number and section ('header' or 'thread')."""
    row_post, row_section = {}, {}
    current_post, section = None, None
    for row in ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        val = str(cell.value).strip() if cell.value is not None else ""
        if POST_HDR.match(val):
            current_post, section = int(POST_HDR.match(val).group(1)), "header"
        elif THREAD_HDR.match(val):
            current_post, section = int(THREAD_HDR.match(val).group(1)), "thread"
        row_post[cell.row] = current_post
        row_section[cell.row] = section
    return row_post, row_section


def collect_links(ws, row_post, row_section):
    """Gather all column-E hyperlink cells, grouped by post."""
    posts = {}
    for row in ws.iter_rows(min_col=5, max_col=5):
        cell = row[0]
        if cell.hyperlink is None:
            continue
        p = row_post.get(cell.row)
        if p is None:
            continue
        target = cell.hyperlink.target or ""
        post_id, comment_id, base = parse_link(target)
        posts.setdefault(p, []).append({
            "cell": cell, "row": cell.row, "section": row_section.get(cell.row),
            "target": target, "post_id": post_id,
            "comment_id": comment_id, "base": base,
        })
    return posts


def authoritative_link(links, post_num, warnings):
    """Decide the correct submission link for one post."""
    header = [l for l in links if l["section"] == "header" and l["comment_id"] is None]
    comments = [l for l in links if l["comment_id"] is not None]

    comment_pids = {l["post_id"] for l in comments if l["post_id"]}
    if len(comment_pids) > 1:
        warnings.append(f"Post {post_num}: comment links span multiple post IDs "
                        f"{sorted(comment_pids)} — fixing skipped, review manually.")
        return None

    if header:
        h = header[0]
        if comment_pids and h["post_id"] not in comment_pids:
            warnings.append(f"Post {post_num}: header link post ID '{h['post_id']}' "
                            f"disagrees with comment links {sorted(comment_pids)} — "
                            f"trusting comment links.")
            c = next(l for l in comments if l["post_id"] in comment_pids)
            stripped = c["target"].rstrip("/")
            return stripped[: stripped.rfind("/")] + "/"
        return h["base"]

    if comments:
        # Derive post link from a comment permalink: strip the comment segment.
        c = comments[0]
        stripped = c["target"].rstrip("/")
        stripped = stripped[: stripped.rfind("/")] + "/"
        return stripped

    warnings.append(f"Post {post_num}: no usable links found — skipped.")
    return None


def fix_sheet(ws):
    row_post, row_section = scan_posts(ws)
    posts = collect_links(ws, row_post, row_section)
    fixes, ok, warnings = [], 0, []

    for post_num in sorted(posts):
        links = posts[post_num]
        correct = authoritative_link(links, post_num, warnings)
        if correct is None:
            continue
        correct_pid, _, _ = parse_link(correct)

        for l in links:
            if l["comment_id"] is not None:        # comment permalink: leave alone
                continue
            if l["section"] != "thread":            # header link is the source of truth
                continue
            if l["post_id"] == correct_pid:
                ok += 1
                continue
            cell = l["cell"]
            cell.hyperlink.target = correct
            if isinstance(cell.value, str) and "reddit.com" in cell.value:
                cell.value = correct
            cell.font = Font(color="0000EE", underline="single")
            fixes.append((cell.row, l["target"], correct))
    return fixes, ok, warnings


def process_file(path):
    wb = load_workbook(path)
    total_fixes, total_ok, all_warnings = 0, 0, []
    for ws in wb.worksheets:
        fixes, ok, warnings = fix_sheet(ws)
        total_fixes += len(fixes)
        total_ok += ok
        all_warnings += warnings
        for row, old, new in fixes:
            old_pid = parse_link(old)[0]
            new_pid = parse_link(new)[0]
            print(f"  row {row}: {old_pid} -> {new_pid}")
    out = os.path.splitext(path)[0] + "_fixed.xlsx"
    wb.save(out)
    print(f"{os.path.basename(path)}: {total_fixes} links fixed, "
          f"{total_ok} already correct -> {os.path.basename(out)}")
    for w in all_warnings:
        print(f"  WARNING: {w}")
    return total_fixes


def main():
    ap = argparse.ArgumentParser(description="Fix wrong OP submission links in ICR sheets")
    ap.add_argument("files", nargs="*", help="xlsx file(s) to fix")
    ap.add_argument("--dir", help="process every .xlsx in this folder")
    args = ap.parse_args()

    targets = list(args.files)
    if args.dir:
        targets += [os.path.join(args.dir, f) for f in sorted(os.listdir(args.dir))
                    if f.endswith(".xlsx") and not f.endswith("_fixed.xlsx")]
    if not targets:
        ap.print_help()
        sys.exit(1)
    for t in targets:
        print(f"\nProcessing {t}")
        process_file(t)


if __name__ == "__main__":
    main()
