#!/usr/bin/env python3
"""Package-level and paragraph-level diff between two .docx files (independent of ASTra).

CLI: docdiff.py before.docx after.docx  -> JSON on stdout
Paragraph identity: body-order list of <w:p> elements (top level and inside tables), each
canonicalised (C14N). "content" view strips w14:paraId / w14:textId / rsid* attributes so that
identifier stamping is reported separately from real content changes.
"""
import difflib, hashlib, json, re, sys, zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STRIP = re.compile(r'\s(?:w14:paraId|w14:textId|w:rsid\w*)="[^"]*"')


def read_parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist() if not n.endswith("/")}


def c14n(data):
    try:
        return etree.tostring(etree.fromstring(data), method="c14n")
    except Exception:
        return data


def paragraphs(xml_bytes):
    root = etree.fromstring(xml_bytes)
    out = []
    for p in root.iter("{%s}p" % W):
        raw = etree.tostring(p, method="c14n").decode("utf8")
        # c14n puts namespace declarations on each element; strip them for stable compare
        raw = re.sub(r'\sxmlns(?::\w+)?="[^"]*"', "", raw)
        text = "".join(t.text or "" for t in p.iter("{%s}t" % W, "{%s}delText" % W))
        out.append({"raw": raw, "content": STRIP.sub("", raw), "text": text})
    return out


def locality(pa, pb, spec):
    """Check that only the intended paragraph changed (content view: ids/rsids stripped).
    spec: target_text, mode in {same, insert, delete}, expect_text, expect_raw (list of substrings)."""
    texts_a = [x["text"] for x in pa]
    cand = [i for i, t in enumerate(texts_a) if t == spec["target_text"]]
    out = {"target_candidates": len(cand)}
    if len(cand) != 1:
        out["target_unique"] = False
        return out
    t = cand[0]
    out["target_unique"] = True
    sm = difflib.SequenceMatcher(a=[x["content"] for x in pa], b=[x["content"] for x in pb], autojunk=False)
    ch_a, ch_b = set(), set()
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag != "equal":
            ch_a.update(range(i1, i2))
            ch_b.update(range(j1, j2))
    mode = spec.get("mode", "same")
    exp_a = {t} if mode in ("same", "delete") else set()
    exp_b = {t} if mode == "same" else ({t + 1} if mode == "insert" else set())
    out["nontarget_changed"] = len(ch_a - exp_a) + len(ch_b - exp_b)
    out["target_changed"] = bool((ch_a & exp_a) or (ch_b & exp_b)) if (exp_a or exp_b) else None
    out["paragraph_delta"] = len(pb) - len(pa)
    # text of every non-target paragraph preserved
    keep_a = [x["text"] for i, x in enumerate(pa) if i != t]
    if mode == "insert":
        keep_a = [x["text"] for x in pa]
        keep_b = [x["text"] for i, x in enumerate(pb) if i != t + 1]
    else:
        keep_b = [x["text"] for i, x in enumerate(pb) if i != t] if mode != "delete" else [x["text"] for x in pb]
    out["nontarget_text_preserved"] = keep_a == keep_b
    pos = t + 1 if mode == "insert" else t
    if mode != "delete" and pos < len(pb):
        tb = pb[pos]
        if "expect_text" in spec:
            out["intended_text_ok"] = tb["text"] == spec["expect_text"]
        if "expect_raw" in spec:
            out["intended_raw_ok"] = all(sub in tb["raw"] for sub in spec["expect_raw"])
    if mode == "delete":
        out["intended_text_ok"] = spec["target_text"] not in [x["text"] for x in pb] or texts_a.count(spec["target_text"]) > 1
    return out


def diff(before, after, spec=None):
    a, b = read_parts(before), read_parts(after)
    names = sorted(set(a) | set(b))
    changed = [n for n in names if a.get(n) != b.get(n)]
    added = [n for n in names if n not in a]
    removed = [n for n in names if n not in b]
    res = {
        "parts_total_before": len(a), "parts_total_after": len(b),
        "parts_changed": [n for n in changed if n in a and n in b],
        "parts_added": added, "parts_removed": removed,
        "bytes_before": sum(len(v) for v in a.values()), "bytes_after": sum(len(v) for v in b.values()),
        "file_bytes_before": __import__("os").path.getsize(before),
        "file_bytes_after": __import__("os").path.getsize(after),
    }
    # byte-level change inside changed parts (size delta + differing prefix/suffix excluded length)
    res["parts_changed_semantic"] = [
        n for n in res["parts_changed"]
        if not n.endswith((".xml", ".rels")) or c14n(a[n]) != c14n(b[n])
    ]
    res["changed_part_bytes_delta"] = {n: len(b[n]) - len(a[n]) for n in res["parts_changed"]}
    da, db = "word/document.xml", "word/document.xml"
    if da in a and db in b:
        pa, pb = paragraphs(a[da]), paragraphs(b[db])
        for view in ("raw", "content"):
            sm = difflib.SequenceMatcher(a=[x[view] for x in pa], b=[x[view] for x in pb], autojunk=False)
            ch_before = ch_after = 0
            ops = []
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag != "equal":
                    ch_before += i2 - i1
                    ch_after += j2 - j1
                    ops.append([tag, i1, i2, j1, j2])
            res[f"paragraphs_{view}_changed_before"] = ch_before
            res[f"paragraphs_{view}_changed_after"] = ch_after
            res[f"opcodes_{view}"] = ops[:20]
        res["paragraphs_before"], res["paragraphs_after"] = len(pa), len(pb)
        if spec:
            res["locality"] = locality(pa, pb, spec)
        res["text_before"] = hashlib.sha256("\n".join(x["text"] for x in pa).encode()).hexdigest()
        res["text_after"] = hashlib.sha256("\n".join(x["text"] for x in pb).encode()).hexdigest()
    return res


if __name__ == "__main__":
    spec = json.loads(sys.argv[3]) if len(sys.argv) > 3 else None
    print(json.dumps(diff(sys.argv[1], sys.argv[2], spec), indent=1))
