#!/usr/bin/env python3
"""Independent oracles O1-O5 (see PROTOCOL.md section 5). None of this is ASTra code.

CLI: oracles.py [--workers N] [--no-lo] file.docx ... | --list list.txt   -> JSON lines on stdout
"""
import concurrent.futures as cf, json, os, posixpath, re, shutil, subprocess, sys, tempfile, zipfile
from lxml import etree

CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def o1_o2_o3(path):
    r = {}
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return {"O1": False, "O1_err": str(e)[:200], "O2": None, "O3": None}
    names = z.namelist()
    bad = z.testzip()
    r["O1"] = bad is None and len(names) == len(set(names))
    if not r["O1"]:
        r["O1_err"] = f"bad_crc={bad} dup={len(names) - len(set(names))}"
    # O2 well-formed
    errs = []
    parts = {}
    for n in names:
        if n.endswith("/"):
            continue
        data = z.read(n)
        parts[n] = data
        if n.endswith((".xml", ".rels")):
            try:
                etree.fromstring(data)
            except Exception as e:
                errs.append(f"{n}: {str(e)[:80]}")
    r["O2"] = not errs
    if errs:
        r["O2_err"] = errs[:5]
    # O3 relationships + content types
    errs = []
    try:
        ct = etree.fromstring(parts["[Content_Types].xml"])
        defaults = {d.get("Extension").lower() for d in ct.findall(CT + "Default")}
        overrides = {o.get("PartName") for o in ct.findall(CT + "Override")}
        for n in parts:
            if n == "[Content_Types].xml":
                continue
            ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
            if "/" + n not in overrides and ext not in defaults:
                errs.append(f"no content type: {n}")
        for n, data in parts.items():
            if not n.endswith(".rels"):
                continue
            base = posixpath.dirname(posixpath.dirname(n))  # folder of the source part
            for rel in etree.fromstring(data).findall(REL + "Relationship"):
                if rel.get("TargetMode") == "External":
                    continue
                t = rel.get("Target")
                full = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(base, t))
                if full not in parts:
                    errs.append(f"{n}: missing target {t}")
    except Exception as e:
        errs.append(f"O3 exception: {str(e)[:100]}")
    r["O3"] = not errs
    if errs:
        r["O3_err"] = errs[:5]
    return r


def o5(path):
    try:
        import docx
        d = docx.Document(path)
        n = sum(1 for _ in d.paragraphs)
        for t in d.tables:
            for row in t.rows:
                for c in row.cells:
                    _ = c.text
        return {"O5": True, "O5_paragraphs": n}
    except Exception as e:
        return {"O5": False, "O5_err": f"{type(e).__name__}: {str(e)[:150]}"}


def o4(path, slot):
    work = tempfile.mkdtemp(prefix=f"astra-lo{slot}-")
    try:
        profile = os.path.join(work, "profile")
        out = os.path.join(work, "out")
        os.makedirs(out)
        p = subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf",
             "--outdir", out, path], capture_output=True, text=True, timeout=600)
        pdf = os.path.join(out, os.path.splitext(os.path.basename(path))[0] + ".pdf")
        if p.returncode != 0 or not os.path.exists(pdf) or os.path.getsize(pdf) < 500:
            return {"O4": False, "O4_err": (p.stderr or p.stdout)[:150], "O4_pages": None}
        info = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
        m = re.search(r"Pages:\s+(\d+)", info)
        return {"O4": True, "O4_pages": int(m.group(1)) if m else None}
    except Exception as e:
        return {"O4": False, "O4_err": f"{type(e).__name__}: {str(e)[:150]}", "O4_pages": None}
    finally:
        shutil.rmtree(work, ignore_errors=True)


def check(args):
    path, slot, use_lo = args
    r = {"file": path}
    r.update(o1_o2_o3(path))
    r.update(o5(path))
    if use_lo:
        r.update(o4(path, slot))
    r["pass_O1_O5"] = all(r.get(k) for k in (["O1", "O2", "O3", "O5"] + (["O4"] if use_lo else [])))
    return r


def main():
    argv = sys.argv[1:]
    workers = 8
    use_lo = True
    files = []
    while argv:
        a = argv.pop(0)
        if a == "--workers":
            workers = int(argv.pop(0))
        elif a == "--no-lo":
            use_lo = False
        elif a == "--list":
            files += [l.strip() for l in open(argv.pop(0)) if l.strip()]
        else:
            files.append(a)
    jobs = [(f, i % workers, use_lo) for i, f in enumerate(files)]
    with cf.ProcessPoolExecutor(max_workers=workers) as ex:
        for res in ex.map(check, jobs, chunksize=1):
            print(json.dumps(res), flush=True)


if __name__ == "__main__":
    main()
