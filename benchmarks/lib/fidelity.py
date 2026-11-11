#!/usr/bin/env python3
"""Independent content-fidelity reader for C1. Compares an expected block list with a .docx.
CLI: fidelity.py out.docx blocks.json -> JSON.
Expected block kinds: h(level,text) p(text) ul(items) ol(items) table(rows) quote(text) code(lines) hr fn(label,text)
"""
import json, sys, zipfile, re
import docx
from docx.oxml.ns import qn

path, blocks_path = sys.argv[1], sys.argv[2]
blocks = json.load(open(blocks_path))
d = docx.Document(path)
items = []  # (kind, style, text, extra)
body = d.element.body
for el in body.iterchildren():
    if el.tag == qn("w:p"):
        p = docx.text.paragraph.Paragraph(el, d)
        ppr = el.find(qn("w:pPr"))
        has_num = ppr is not None and ppr.find(qn("w:numPr")) is not None
        items.append(("p", p.style.name if p.style is not None else "", p.text, {"num": has_num}))
    elif el.tag == qn("w:tbl"):
        t = docx.table.Table(el, d)
        rows = [[c.text for c in r.cells] for r in t.rows]
        items.append(("t", "", "", {"rows": rows}))

def norm(s):
    return re.sub(r"\s+", " ", s).strip()

cursor = 0
def find(pred):
    global cursor
    for i in range(cursor, len(items)):
        if pred(items[i]):
            cursor = i + 1
            return items[i]
    return None

res = {"expected": 0, "found_in_order": 0, "attr_ok": 0, "attr_checked": 0, "missing": []}
fn_text = ""
with zipfile.ZipFile(path) as z:
    if "word/footnotes.xml" in z.namelist():
        fn_text = norm(re.sub(r"<[^>]+>", " ", z.read("word/footnotes.xml").decode("utf8", "replace")))

def want(kind, ok_text, attr=None, label=""):
    res["expected"] += 1
    it = find(ok_text)
    if it is None:
        res["missing"].append(f"{kind}:{label[:40]}")
        return
    res["found_in_order"] += 1
    if attr is not None:
        res["attr_checked"] += 1
        res["attr_ok"] += 1 if attr(it) else 0

for b in blocks:
    k = b["kind"]
    if k == "h":
        want(k, lambda it, b=b: it[0] == "p" and norm(it[2]) == norm(b["text"]),
             lambda it, b=b: it[1].replace(" ", "").lower() in (f"heading{b['level']}",), b["text"])
    elif k == "p":
        want(k, lambda it, b=b: it[0] == "p" and norm(it[2]) == norm(b["text"]), None, b["text"])
    elif k in ("ul", "ol"):
        for text in b["items"]:
            want(k, lambda it, t=text: it[0] == "p" and norm(it[2]) == norm(t), lambda it: it[3]["num"] or "List" in it[1], text)
    elif k == "table":
        shape = (len(b["rows"]), len(b["rows"][0]))
        want(k, lambda it, b=b: it[0] == "t" and [[norm(c) for c in r] for r in it[3]["rows"]] == [[norm(c) for c in r] for r in b["rows"]],
             lambda it, shape=shape: (len(it[3]["rows"]), len(it[3]["rows"][0])) == shape, "table")
    elif k == "quote":
        want(k, lambda it, b=b: it[0] == "p" and norm(it[2]) == norm(b["text"]), lambda it: "quote" in it[1].lower(), b["text"])
    elif k == "code":
        # a fenced code block is written as ONE paragraph with line breaks
        want(k, lambda it, b=b: it[0] == "p" and all(norm(l) in norm(it[2]) for l in b["lines"]), None, b["lines"][0])
    elif k == "fn":
        res["expected"] += 1
        if norm(b["text"]) in fn_text:
            res["found_in_order"] += 1
        else:
            res["missing"].append(f"fn:{b['label']}")
# presence anywhere (order-independent) for text-bearing blocks, as a second view
all_text = norm(" ".join([i[2] for i in items if i[0] == "p"] + [c for i in items if i[0] == "t" for r in i[3]["rows"] for c in r]))
res["found_anywhere"] = sum(1 for b in blocks if b["kind"] in ("h", "p", "quote") and norm(b["text"]) in all_text) + sum(
    1 for b in blocks if b["kind"] in ("ul", "ol") for t in b["items"] if norm(t) in all_text) + sum(
    1 for b in blocks if b["kind"] == "code" and all(norm(l) in all_text for l in b["lines"]))
res["expected_text_blocks"] = sum(1 for b in blocks if b["kind"] in ("h", "p", "quote", "code")) + sum(len(b["items"]) for b in blocks if b["kind"] in ("ul", "ol"))
res["unexpected_paragraphs"] = max(0, sum(1 for i in items if i[0] == "p" and norm(i[2])) - sum(
    1 if b["kind"] in ("h", "p", "quote") else len(b.get("items", b.get("lines", []))) if b["kind"] in ("ul", "ol", "code") else 0 for b in blocks)
    - sum(len(r) for b in blocks if b["kind"] == "table" for r in b["rows"]) * 0)
res["output_blocks"] = len(items)
print(json.dumps(res))
