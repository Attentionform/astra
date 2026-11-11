#!/usr/bin/env python3
"""Deterministic corpus generator for the ASTra benchmark suite.

Writers: python-docx (independent producer), LibreOffice (HTML -> DOCX), lxml injection
for tracked changes, bookmarks and hyperlinks. Seeds are fixed. No private data.

Usage: python3 corpus/generate.py [out_dir]
Writes <out_dir>/*.docx and <out_dir>/manifest.json (default: corpus/files).
"""
import hashlib, io, json, os, random, shutil, subprocess, sys, tempfile, zipfile
from datetime import datetime, timezone

import docx
from docx.enum.text import WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PIL import Image

FIXED_DT = datetime(2026, 1, 1, 0, 0, 0)
PARAS_PER_PAGE = 5

VOCAB = (
    "account agreement analysis annual approach asset assumption audit balance baseline benefit board "
    "budget capacity capital change client clause committee compliance contract control cost customer "
    "data deadline delivery demand design determine document driver effect efficiency employee estimate "
    "evidence execution expense factor forecast framework funding goal governance growth impact income "
    "indicator initiative insight investment issue launch liability liquidity margin market measure "
    "milestone objective obligation operation option outcome partner performance period pipeline plan "
    "policy portfolio priority process procurement product profit project proposal quality quarter "
    "recommendation record reduction regulation report requirement reserve resource result review revenue "
    "risk roadmap scenario schedule scope segment service signal stakeholder standard strategy summary "
    "supplier target team term threshold timeline trend update utilization valuation variance vendor "
    "volume workflow"
).split()


def sentence(rng, lo=8, hi=18):
    n = rng.randint(lo, hi)
    words = [rng.choice(VOCAB) for _ in range(n)]
    words[0] = words[0].capitalize()
    return " ".join(words) + "."


def paragraph_text(rng, words=60):
    out, count = [], 0
    while count < words:
        s = sentence(rng)
        out.append(s)
        count += len(s.split())
    return " ".join(out)


def new_doc():
    d = docx.Document()
    cp = d.core_properties
    cp.created = cp.modified = cp.last_printed = FIXED_DT
    cp.author = "ASTra benchmark corpus"
    cp.last_modified_by = "ASTra benchmark corpus"
    cp.title = ""
    return d


def add_runs(rng, p, text):
    """Split text into runs with occasional bold/italic to give realistic run structure."""
    words = text.split(" ")
    i = 0
    while i < len(words):
        n = rng.randint(4, 14)
        chunk = " ".join(words[i:i + n]) + (" " if i + n < len(words) else "")
        r = p.add_run(chunk)
        roll = rng.random()
        if roll < 0.08:
            r.bold = True
        elif roll < 0.14:
            r.italic = True
        i += n


def png_bytes(rng, w=400, h=200):
    c1 = tuple(rng.randint(30, 220) for _ in range(3))
    c2 = tuple(rng.randint(30, 220) for _ in range(3))
    img = Image.new("RGB", (w, h))
    px = img.load()
    for x in range(w):
        t = x / (w - 1)
        col = tuple(int(c1[k] * (1 - t) + c2[k] * t) for k in range(3))
        for y in range(h):
            px[x, y] = col
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf


def add_table(rng, d, rows, cols, header=True, merge=False):
    t = d.add_table(rows=rows, cols=cols)
    t.style = "Table Grid"
    for c in range(cols):
        cell = t.cell(0, c)
        cell.text = ""
        run = cell.paragraphs[0].add_run(rng.choice(VOCAB).capitalize())
        run.bold = header
    for r in range(1, rows):
        for c in range(cols):
            t.cell(r, c).text = f"{rng.choice(VOCAB)} {rng.randint(1, 9999)}"
    if merge and rows > 3 and cols > 2:
        t.cell(1, 0).merge(t.cell(1, 1))
    return t


def prose(pages, seed):
    rng = random.Random(seed)
    d = new_doc()
    d.add_heading(f"Report {seed}", 0)
    total = pages * PARAS_PER_PAGE
    for i in range(total):
        if i % 12 == 0:
            d.add_heading(sentence(rng, 3, 6).rstrip("."), 1)
        elif i % 6 == 0:
            d.add_heading(sentence(rng, 2, 5).rstrip("."), 2)
        p = d.add_paragraph()
        add_runs(rng, p, paragraph_text(rng))
    return d


def tables(pages, seed, cols=5):
    rng = random.Random(seed)
    d = new_doc()
    d.add_heading(f"Tables {seed}", 0)
    for i in range(max(1, round(pages * 2.5))):
        d.add_heading(sentence(rng, 2, 5).rstrip("."), 2)
        add_table(rng, d, rows=rng.randint(6, 12), cols=cols, merge=(i % 3 == 0))
        p = d.add_paragraph()
        add_runs(rng, p, paragraph_text(rng, 30))
    return d


def lists(pages, seed):
    rng = random.Random(seed)
    d = new_doc()
    d.add_heading(f"Lists {seed}", 0)
    for i in range(max(1, round(pages * 1.4))):
        d.add_heading(sentence(rng, 2, 5).rstrip("."), 2)
        for j in range(rng.randint(5, 9)):
            d.add_paragraph(sentence(rng), style="List Bullet" if j % 4 else "List Bullet 2")
        for j in range(rng.randint(3, 6)):
            d.add_paragraph(sentence(rng), style="List Number")
        p = d.add_paragraph()
        add_runs(rng, p, paragraph_text(rng, 40))
    return d


def mixed(pages, seed):
    rng = random.Random(seed)
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = f"Benchmark corpus document {seed}"
    fp = d.sections[0].footer.paragraphs[0]
    fp.text = "Page "
    run = fp.add_run()
    for kind, txt in (("begin", None), (None, " PAGE "), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = txt
        run._r.append(el)
    d.add_heading(f"Mixed {seed}", 0)
    for i in range(pages):
        d.add_heading(sentence(rng, 2, 5).rstrip("."), 1)
        for _ in range(2):
            p = d.add_paragraph()
            add_runs(rng, p, paragraph_text(rng))
        add_table(rng, d, rows=rng.randint(4, 8), cols=rng.randint(3, 6), merge=(i % 2 == 0))
        for j in range(4):
            d.add_paragraph(sentence(rng), style="List Bullet")
        if i % 2 == 0:
            d.add_picture(png_bytes(rng), width=Inches(4))
        d.add_paragraph(sentence(rng, 4, 8), style="Caption")
    return d


def with_comments(pages, seed):
    rng = random.Random(seed)
    d = prose(pages, seed)
    paras = [p for p in d.paragraphs if p.runs and len(p.text) > 80]
    for k, p in enumerate(paras[:: max(1, len(paras) // 12)][:12]):
        d.add_comment(runs=p.runs[0], text=sentence(rng, 6, 12), author="Reviewer", initials="RV")
    return d


def inject_tracked(d, seed):
    """Wrap some runs in w:ins / w:del. Pure lxml; deterministic."""
    rng = random.Random(seed)
    rev = 1
    paras = [p for p in d.paragraphs if len(p.runs) >= 3]
    for p in paras[:: max(1, len(paras) // 16)][:16]:
        runs = p.runs
        target = runs[1]._r
        wrapper = OxmlElement("w:ins" if rev % 2 else "w:del")
        wrapper.set(qn("w:id"), str(rev))
        wrapper.set(qn("w:author"), "Reviewer")
        wrapper.set(qn("w:date"), "2026-01-01T00:00:00Z")
        parent = target.getparent()
        parent.replace(target, wrapper)
        wrapper.append(target)
        if rev % 2 == 0:
            for t in target.findall(qn("w:t")):
                t.tag = qn("w:delText")
        rev += 1
    return d


def inject_links_bookmarks(d, seed):
    rng = random.Random(seed)
    part = d.part
    bid = 1
    paras = [p for p in d.paragraphs if p.runs and len(p.text) > 60]
    for k, p in enumerate(paras[:: max(1, len(paras) // 20)][:20]):
        # bookmark around the paragraph content
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(bid))
        start.set(qn("w:name"), f"bm_{seed}_{bid}")
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(bid))
        pel = p._p
        first_run = p.runs[0]._r
        first_run.addprevious(start)
        pel.append(end)
        bid += 1
        if k % 2 == 0:
            rid = part.relate_to(f"https://example.org/doc/{seed}/{k}", RT.HYPERLINK, is_external=True)
            h = OxmlElement("w:hyperlink")
            h.set(qn("r:id"), rid)
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.text = " see reference "
            r.append(t)
            h.append(r)
            pel.append(h)
    return d


def heavy_runs(pages, seed):
    """Many small formatted runs per paragraph (stress for run splitting)."""
    rng = random.Random(seed)
    d = new_doc()
    d.add_heading(f"Runs {seed}", 0)
    for i in range(pages * PARAS_PER_PAGE):
        p = d.add_paragraph()
        for _ in range(rng.randint(10, 24)):
            r = p.add_run(" ".join(rng.choice(VOCAB) for _ in range(rng.randint(1, 4))) + " ")
            r.bold = rng.random() < 0.3
            r.italic = rng.random() < 0.2
            r.font.size = Pt(rng.choice([10, 11, 12]))
    return d


# ---------- LibreOffice writer (HTML -> DOCX) ----------

def lo_html(kind, pages, seed):
    rng = random.Random(seed)
    h = ["<html><head><meta charset='utf-8'><title>LO corpus</title></head><body>"]
    h.append(f"<h1>LibreOffice {kind} {seed}</h1>")
    for i in range(pages * (2 if kind != "prose" else 3)):
        h.append(f"<h2>{sentence(rng, 2, 5).rstrip('.')}</h2>")
        if kind in ("prose", "mixed"):
            for _ in range(3 if kind == "prose" else 2):
                h.append(f"<p>{paragraph_text(rng).replace('.', '.', 1)} <b>{rng.choice(VOCAB)}</b> <i>{rng.choice(VOCAB)}</i></p>")
        if kind in ("tables", "mixed"):
            cols = rng.randint(3, 6)
            h.append("<table border='1' cellspacing='0'>")
            for r in range(rng.randint(5, 9)):
                tag = "th" if r == 0 else "td"
                h.append("<tr>" + "".join(f"<{tag}>{rng.choice(VOCAB)} {rng.randint(1, 999)}</{tag}>" for _ in range(cols)) + "</tr>")
            h.append("</table>")
        if kind in ("lists", "mixed"):
            h.append("<ul>" + "".join(f"<li>{sentence(rng)}</li>" for _ in range(rng.randint(4, 7))) + "</ul>")
            h.append("<ol>" + "".join(f"<li>{sentence(rng)}</li>" for _ in range(rng.randint(3, 5))) + "</ol>")
    h.append("</body></html>")
    return "\n".join(h)


def lo_convert(html, out_path, workdir):
    src = os.path.join(workdir, "src.html")
    with open(src, "w", encoding="utf-8") as f:
        f.write(html)
    profile = os.path.join(workdir, "lo-profile")
    outdir = os.path.join(workdir, "lo-out")
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(
        ["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to",
         "docx:MS Word 2007 XML", "--outdir", outdir, src],
        check=True, capture_output=True, timeout=300,
    )
    shutil.move(os.path.join(outdir, "src.docx"), out_path)


# ---------- determinism + manifest ----------

def rezip_fixed(path):
    """Rewrite the zip with fixed timestamps and sorted order so bytes are stable."""
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}
    order = ["[Content_Types].xml"] + sorted(n for n in names if n != "[Content_Types].xml")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
        for n in order:
            zi = zipfile.ZipInfo(n, date_time=(2026, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zout.writestr(zi, data[n])
    os.replace(tmp, path)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def count_features(path):
    d = docx.Document(path)
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf8", "replace")
        images = sum(1 for n in z.namelist() if n.startswith("word/media/"))
    return {
        "paragraphs": len(d.paragraphs),
        "tables": len(d.tables),
        "images": images,
        "tracked_insertions": xml.count("<w:ins "),
        "tracked_deletions": xml.count("<w:del "),
        "comments": len(list(d.comments)) if hasattr(d, "comments") else 0,
        "hyperlinks": xml.count("<w:hyperlink"),
        "bookmarks": xml.count("<w:bookmarkStart"),
    }


def measure_pages(path, workdir):
    profile = os.path.join(workdir, "lo-profile-pages")
    outdir = os.path.join(workdir, "pdf")
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(
        ["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf",
         "--outdir", outdir, path],
        check=True, capture_output=True, timeout=600,
    )
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    for line in out.splitlines():
        if line.startswith("Pages:"):
            return int(line.split()[1])
    return None


SPECS = []
def spec(name, writer, kind, pages, seed, build):
    SPECS.append(dict(name=name, writer=writer, kind=kind, target_pages=pages, seed=seed, build=build))

for pg in (1, 2, 5, 10, 25, 50, 100, 200):
    spec(f"pd-prose-{pg:03d}p", "python-docx", "prose", pg, 1000 + pg, lambda pg=pg: prose(pg, 1000 + pg))
for s in (2, 3):
    spec(f"pd-prose-003p-s{s}", "python-docx", "prose", 3, 1003 * s, lambda s=s: prose(3, 1003 * s))
for pg in (1, 5, 25, 100):
    spec(f"pd-tables-{pg:03d}p", "python-docx", "tables", pg, 2000 + pg, lambda pg=pg: tables(pg, 2000 + pg))
spec("pd-tables-wide-010p", "python-docx", "tables-wide", 10, 2210, lambda: tables(10, 2210, cols=12))
for pg in (1, 5, 25, 100):
    spec(f"pd-lists-{pg:03d}p", "python-docx", "lists", pg, 3000 + pg, lambda pg=pg: lists(pg, 3000 + pg))
for pg in (2, 10, 50, 100, 200):
    spec(f"pd-mixed-{pg:03d}p", "python-docx", "mixed", pg, 4000 + pg, lambda pg=pg: mixed(pg, 4000 + pg))
for pg in (5, 25):
    spec(f"pd-comments-{pg:03d}p", "python-docx", "comments", pg, 5000 + pg, lambda pg=pg: with_comments(pg, 5000 + pg))
for pg in (5, 25):
    spec(f"pd-tracked-{pg:03d}p", "python-docx+lxml", "tracked", pg, 6000 + pg, lambda pg=pg: inject_tracked(prose(pg, 6000 + pg), 6000 + pg))
spec("pd-links-bookmarks-005p", "python-docx+lxml", "links-bookmarks", 5, 7005, lambda: inject_links_bookmarks(prose(5, 7005), 7005))
spec("pd-runs-010p", "python-docx", "heavy-runs", 10, 8010, lambda: heavy_runs(10, 8010))
for pg in (1, 5, 25, 100):
    spec(f"lo-prose-{pg:03d}p", "libreoffice", "prose", pg, 9000 + pg, ("lo", "prose", pg, 9000 + pg))
for pg in (1, 5, 25):
    spec(f"lo-tables-{pg:03d}p", "libreoffice", "tables", pg, 9100 + pg, ("lo", "tables", pg, 9100 + pg))
for pg in (1, 5):
    spec(f"lo-lists-{pg:03d}p", "libreoffice", "lists", pg, 9200 + pg, ("lo", "lists", pg, 9200 + pg))
spec("lo-mixed-010p", "libreoffice", "mixed", 10, 9310, ("lo", "mixed", 10, 9310))


def main():
    out = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "files"))
    os.makedirs(out, exist_ok=True)
    work = tempfile.mkdtemp(prefix="astra-corpus-")
    manifest = {
        "generated_for": "ASTra benchmark corpus v1",
        "python_docx": docx.__version__ if hasattr(docx, "__version__") else "unknown",
        "libreoffice": subprocess.run(["soffice", "--version"], capture_output=True, text=True).stdout.strip(),
        "documents": [],
    }
    for s in SPECS:
        path = os.path.join(out, s["name"] + ".docx")
        b = s["build"]
        if isinstance(b, tuple):
            _, kind, pg, seed = b
            lo_convert(lo_html(kind, pg, seed), path, work)
        else:
            b().save(path)
            rezip_fixed(path)
        feats = count_features(path)
        pages = measure_pages(path, work)
        manifest["documents"].append({
            "file": s["name"] + ".docx", "writer": s["writer"], "kind": s["kind"],
            "target_pages": s["target_pages"], "measured_pages": pages, "seed": s["seed"],
            "bytes": os.path.getsize(path), "sha256": sha256(path), **feats,
        })
        print(f"{s['name']:28s} pages={pages} bytes={os.path.getsize(path)}", flush=True)
    with open(os.path.join(os.path.dirname(out), "manifest.json"), "w") as f:
        json.dump(manifest, f, indent=2)
    shutil.rmtree(work, ignore_errors=True)
    print(f"wrote {len(manifest['documents'])} documents to {out}")


if __name__ == "__main__":
    main()
