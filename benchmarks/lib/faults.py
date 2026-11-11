#!/usr/bin/env python3
"""Fault injection for E3. 60 invalid-by-construction mutations in six categories + 20 valid controls.
Usage: faults.py <corpus_dir> <out_dir>   -> writes <out_dir>/*.docx and <out_dir>/faults.json
Each fault is an XML/package mutation of a valid base document; the label is from the ECMA-376 rule
named in `rule`. kind: schema (XSD-level), semantic (cross-reference; not XSD), package (OPC), wellformed.
"""
import copy, json, os, re, sys, zipfile
from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
W = "{%s}" % NS["w"]
BASES = ["pd-mixed-010p", "lo-mixed-010p", "pd-comments-005p"]


def q(tag):
    p, n = tag.split(":")
    return "{%s}%s" % (NS[p], n)


class Pkg:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.order = z.namelist()
        self.parts = {n: z.read(n) for n in self.order}
        self._trees = {}

    def tree(self, name):
        if name not in self._trees:
            self._trees[name] = etree.fromstring(self.parts[name])
        return self._trees[name]

    def commit(self):
        for n, t in self._trees.items():
            self.parts[n] = etree.tostring(t, xml_declaration=True, encoding="UTF-8", standalone=True)

    def save(self, path):
        self.commit()
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                if n in self.parts:
                    z.writestr(n, self.parts[n])

    def x(self, xp, part="word/document.xml"):
        return self.tree(part).xpath(xp, namespaces=NS)


def first(p, xp, part="word/document.xml", nth=0):
    r = p.x(xp, part)
    return r[nth] if len(r) > nth else None


def move_before(a, b):
    b.getparent().insert(b.getparent().index(b), a)


def E(tag, **attrs):
    el = etree.Element(q("w:" + tag))
    for k, v in attrs.items():
        el.set(q("w:" + k) if ":" not in k else q(k), v)
    return el


FAULTS = []
def fault(fid, cat, kind, title, rule):
    def deco(fn):
        FAULTS.append(dict(id=fid, category=cat, kind=kind, title=title, rule=rule, fn=fn))
        return fn
    return deco


# ---- A. element order ------------------------------------------------------
@fault("A01", "order", "schema", "pPr: jc before pStyle", "CT_PPr sequence order")
def _(p):
    ppr = first(p, "//w:pPr[w:pStyle]")
    if ppr is None: return False
    ppr.insert(0, E("jc", val="center")); return True

@fault("A02", "order", "schema", "pPr: spacing before keepNext", "CT_PPr sequence order")
def _(p):
    ppr = first(p, "//w:pPr[w:keepNext]")
    if ppr is None:
        ppr = first(p, "//w:pPr")
        if ppr is None: return False
        ppr.insert(0, E("keepNext"))
    ppr.insert(0, E("spacing", after="120")); return True

@fault("A03", "order", "schema", "tblPr: tblW before tblStyle", "CT_TblPr sequence order")
def _(p):
    t = first(p, "//w:tblPr[w:tblStyle]")
    if t is None: return False
    t.insert(0, E("tblW", w="5000", type="dxa")); return True

@fault("A04", "order", "schema", "tcPr: shd before tcW", "CT_TcPr sequence order")
def _(p):
    t = first(p, "//w:tcPr[w:tcW]")
    if t is None: return False
    t.insert(0, E("shd", val="clear", fill="DDDDDD")); return True

@fault("A05", "order", "schema", "sectPr: pgMar before pgSz", "CT_SectPr sequence order")
def _(p):
    s = first(p, "//w:sectPr[w:pgSz][w:pgMar]")
    if s is None: return False
    move_before(s.find(q("w:pgMar")), s.find(q("w:pgSz"))); return True

@fault("A06", "order", "schema", "body: sectPr is first child", "CT_Body: sectPr is last")
def _(p):
    b = first(p, "//w:body"); s = b.find(q("w:sectPr"))
    if s is None: return False
    b.remove(s); b.insert(0, s); return True

@fault("A07", "order", "schema", "p: pPr after a run", "CT_P: pPr first")
def _(p):
    para = first(p, "//w:p[w:pPr][w:r]")
    if para is None: return False
    ppr = para.find(q("w:pPr")); para.remove(ppr); para.append(ppr); return True

@fault("A08", "order", "schema", "tbl: tblGrid before tblPr", "CT_Tbl sequence order")
def _(p):
    t = first(p, "//w:tbl[w:tblPr][w:tblGrid]")
    if t is None: return False
    move_before(t.find(q("w:tblGrid")), t.find(q("w:tblPr"))); return True

@fault("A09", "order", "schema", "tr: tc before trPr", "CT_Row order")
def _(p):
    tr = first(p, "//w:tr[w:trPr]")
    if tr is None:
        tr = first(p, "//w:tr")
        if tr is None: return False
        tr.insert(0, E("trPr")); 
    trpr = tr.find(q("w:trPr")); tr.remove(trpr); tr.append(trpr); return True

@fault("A10", "order", "schema", "pPr: ind before spacing", "CT_PPr sequence order")
def _(p):
    ppr = first(p, "//w:pPr[w:spacing]")
    if ppr is None:
        ppr = first(p, "//w:pPr")
        if ppr is None: return False
        ppr.append(E("spacing", after="100"))
    sp = ppr.find(q("w:spacing")); ind = E("ind", left="200"); move_before(ind, sp); return True


# ---- B. attribute values -----------------------------------------------------
def with_attr_target(p, xp, attr, value, make=None):
    el = first(p, xp)
    if el is None:
        return False
    el.set(q("w:" + attr), value); return True

def ensure_rpr_child(p, tag, **attrs):
    r = first(p, "//w:r[w:t]")
    if r is None: return None
    rpr = r.find(q("w:rPr"))
    if rpr is None:
        rpr = E("rPr"); r.insert(0, rpr)
    el = E(tag, **attrs); rpr.append(el); return el

@fault("B01", "attribute", "schema", "jc val=middle", "ST_Jc enumeration")
def _(p):
    ppr = first(p, "//w:pPr")
    if ppr is None: return False
    ppr.append(E("jc", val="middle")); return True

@fault("B02", "attribute", "schema", "sz val=abc", "ST_HpsMeasure numeric")
def _(p): return ensure_rpr_child(p, "sz", val="abc") is not None

@fault("B03", "attribute", "schema", "szCs val=twelve", "ST_HpsMeasure numeric")
def _(p): return ensure_rpr_child(p, "szCs", val="twelve") is not None

@fault("B04", "attribute", "schema", "color val=GGGGGG", "ST_HexColor pattern")
def _(p): return ensure_rpr_child(p, "color", val="GGGGGG") is not None

@fault("B05", "attribute", "schema", "shd fill=notahex", "ST_HexColor pattern")
def _(p):
    tc = first(p, "//w:tcPr")
    if tc is None: return False
    tc.append(E("shd", val="clear", fill="notahex")); return True

@fault("B06", "attribute", "schema", "u val=wavy2", "ST_Underline enumeration")
def _(p): return ensure_rpr_child(p, "u", val="wavy2") is not None

@fault("B07", "attribute", "schema", "spacing line=abc", "ST_SignedTwipsMeasure numeric")
def _(p):
    ppr = first(p, "//w:pPr")
    if ppr is None: return False
    ppr.append(E("spacing", line="abc")); return True

@fault("B08", "attribute", "schema", "ind left=1e5x", "ST_SignedTwipsMeasure numeric")
def _(p):
    ppr = first(p, "//w:pPr")
    if ppr is None: return False
    ppr.append(E("ind", left="1e5x")); return True

@fault("B09", "attribute", "schema", "tblW type=fraction", "ST_TblWidth enumeration")
def _(p):
    t = first(p, "//w:tblW")
    if t is None: return False
    t.set(q("w:type"), "fraction"); return True

@fault("B10", "attribute", "schema", "highlight val=purple", "ST_HighlightColor enumeration")
def _(p): return ensure_rpr_child(p, "highlight", val="purple") is not None


# ---- C. unknown content -------------------------------------------------------
@fault("C01", "unknown", "schema", "unknown element in p", "CT_P content model")
def _(p):
    para = first(p, "//w:p")
    para.append(E("bogus")); return True

@fault("C02", "unknown", "schema", "unknown element in rPr", "CT_RPr content model")
def _(p): return ensure_rpr_child(p, "bogus") is not None

@fault("C03", "unknown", "schema", "unknown attribute on p", "CT_P attributes")
def _(p):
    para = first(p, "//w:p"); para.set(q("w:foo"), "1"); return True

@fault("C04", "unknown", "schema", "unknown element in body", "CT_Body content model")
def _(p):
    b = first(p, "//w:body"); b.insert(0, E("bogus")); return True

@fault("C05", "unknown", "schema", "unknown element in tc", "CT_Tc content model")
def _(p):
    tc = first(p, "//w:tc")
    if tc is None: return False
    tc.insert(1, E("bogus")); return True

@fault("C06", "unknown", "schema", "t xml:space=bad", "xml:space enumeration")
def _(p):
    t = first(p, "//w:t"); t.set("{http://www.w3.org/XML/1998/namespace}space", "bad"); return True

@fault("C07", "unknown", "schema", "foreign-namespace element in p", "CT_P content model")
def _(p):
    para = first(p, "//w:p"); para.append(etree.Element("{http://example.org/x}foo")); return True

@fault("C08", "unknown", "schema", "unknown child of pPr", "CT_PPr content model")
def _(p):
    ppr = first(p, "//w:pPr")
    if ppr is None: return False
    ppr.append(E("notAProp")); return True

@fault("C09", "unknown", "schema", "unknown child of tbl", "CT_Tbl content model")
def _(p):
    t = first(p, "//w:tbl")
    if t is None: return False
    t.append(E("bogus")); return True

@fault("C10", "unknown", "schema", "unknown child of r", "CT_R content model")
def _(p):
    r = first(p, "//w:r[w:t]"); r.append(E("bogus")); return True


# ---- D. missing required content ----------------------------------------------
@fault("D01", "missing", "schema", "tc without any paragraph", "CT_Tc requires a block-level element")
def _(p):
    tc = first(p, "//w:tc")
    if tc is None: return False
    for ch in list(tc):
        if ch.tag != q("w:tcPr"): tc.remove(ch)
    return True

@fault("D02", "missing", "schema", "tbl without tblGrid", "CT_Tbl requires tblGrid")
def _(p):
    t = first(p, "//w:tbl[w:tblGrid]")
    if t is None: return False
    t.remove(t.find(q("w:tblGrid"))); return True

@fault("D03", "missing", "schema", "tbl without rows", "CT_Tbl requires at least one tr")
def _(p):
    t = first(p, "//w:tbl")
    if t is None: return False
    for tr in t.findall(q("w:tr")): t.remove(tr)
    return True

@fault("D04", "missing", "schema", "bookmarkStart without id", "CT_BookmarkStart requires w:id")
def _(p):
    para = first(p, "//w:p"); b = E("bookmarkStart", name="bm_x"); para.append(b); return True

@fault("D05", "missing", "schema", "bookmarkStart without name", "CT_Bookmark requires w:name")
def _(p):
    para = first(p, "//w:p"); b = E("bookmarkStart", id="990"); para.append(b); return True

@fault("D06", "missing", "schema", "commentReference without id", "CT_Markup requires w:id")
def _(p):
    r = first(p, "//w:r[w:t]"); r.append(E("commentReference")); return True

@fault("D07", "missing", "schema", "ins without author", "CT_TrackChange requires w:author")
def _(p):
    para = first(p, "//w:p"); ins = E("ins", id="991"); para.append(ins); return True

@fault("D08", "missing", "schema", "fldChar without type", "CT_FldChar requires fldCharType")
def _(p):
    r = first(p, "//w:r[w:t]"); r.append(E("fldChar")); return True

@fault("D09", "missing", "schema", "del without id", "CT_TrackChange requires w:id")
def _(p):
    para = first(p, "//w:p"); para.append(E("del", author="x", date="2026-01-01T00:00:00Z")); return True

@fault("D10", "missing", "schema", "drawing without inline or anchor", "CT_Drawing content model")
def _(p):
    r = first(p, "//w:r[w:t]"); r.append(E("drawing")); return True


# ---- E. identifiers and cross references ---------------------------------------
@fault("E01", "identifier", "schema", "docPr id=abc", "ST_DrawingElementId unsignedInt")
def _(p):
    d = first(p, "//wp:docPr")
    if d is None: return False
    d.set("id", "abc"); return True

@fault("E02", "identifier", "semantic", "duplicate docPr ids", "unique drawing object ids")
def _(p):
    ds = p.x("//wp:docPr")
    if len(ds) < 2: return False
    ds[1].set("id", ds[0].get("id")); return True

@fault("E03", "identifier", "semantic", "duplicate bookmark ids", "unique bookmark ids")
def _(p):
    para = first(p, "//w:p")
    for n in ("a", "b"):
        para.append(E("bookmarkStart", id="777", name="dup_" + n))
        para.append(E("bookmarkEnd", id="777"))
    return True

@fault("E04", "identifier", "semantic", "bookmarkEnd without start", "bookmark pairing")
def _(p):
    first(p, "//w:p").append(E("bookmarkEnd", id="888")); return True

@fault("E05", "identifier", "semantic", "commentRangeStart refers to no comment", "comment id exists in comments part")
def _(p):
    para = first(p, "//w:p"); para.insert(0, E("commentRangeStart", id="4242")); para.append(E("commentRangeEnd", id="4242")); return True

@fault("E06", "identifier", "semantic", "numId refers to missing num", "numbering reference")
def _(p):
    ppr = first(p, "//w:pPr")
    if ppr is None: return False
    num = E("numPr"); num.append(E("ilvl", val="0")); num.append(E("numId", val="9999")); ppr.append(num); return True

@fault("E07", "identifier", "semantic", "pStyle refers to missing style", "style reference")
def _(p):
    ppr = first(p, "//w:pPr")
    if ppr is None: return False
    for old in ppr.findall(q("w:pStyle")): ppr.remove(old)
    ppr.insert(0, E("pStyle", val="NoSuchStyle123")); return True

@fault("E08", "identifier", "semantic", "footnoteReference to missing footnote", "footnote id exists")
def _(p):
    r = first(p, "//w:r[w:t]"); r.append(E("footnoteReference", id="9999")); return True

@fault("E09", "identifier", "schema", "ins id not a number", "ST_DecimalNumber")
def _(p):
    para = first(p, "//w:p"); para.append(E("ins", id="x1", author="a", date="2026-01-01T00:00:00Z")); return True

@fault("E10", "identifier", "schema", "bookmarkStart id not a number", "ST_DecimalNumber")
def _(p):
    para = first(p, "//w:p"); para.append(E("bookmarkStart", id="one", name="bad_id")); return True


# ---- F. package / relationships -------------------------------------------------
def rels(p, name="word/_rels/document.xml.rels"):
    return p.tree(name)

@fault("F01", "package", "package", "relationship target part deleted", "OPC relationship target exists")
def _(p):
    media = [n for n in p.parts if n.startswith("word/media/")]
    if not media: return False
    del p.parts[media[0]]; return True

@fault("F02", "package", "package", "r:embed to missing relationship", "relationship id exists")
def _(p):
    blip = first(p, "//a:blip")
    if blip is None: return False
    blip.set("{%s}embed" % NS["r"], "rId9999"); return True

@fault("F03", "package", "package", "styles override removed from content types", "content type for every part")
def _(p):
    ct = p.tree("[Content_Types].xml")
    for o in ct:
        if o.get("PartName") == "/word/styles.xml":
            ct.remove(o); return True
    return False

@fault("F04", "package", "package", "styles part deleted, relationship kept", "OPC relationship target exists")
def _(p):
    if "word/styles.xml" not in p.parts: return False
    del p.parts["word/styles.xml"]; return True

@fault("F05", "package", "package", "duplicate relationship id", "unique relationship ids")
def _(p):
    rel = rels(p); items = list(rel)
    if len(items) < 2: return False
    items[1].set("Id", items[0].get("Id")); return True

@fault("F06", "package", "package", "hyperlink r:id missing from relationships", "relationship id exists")
def _(p):
    para = first(p, "//w:p")
    h = E("hyperlink"); h.set("{%s}id" % NS["r"], "rId8888"); rr = E("r"); t = E("t"); t.text = "x"; rr.append(t); h.append(rr); para.append(h); return True

@fault("F07", "package", "package", "header part deleted, relationship kept", "OPC relationship target exists")
def _(p):
    hd = [n for n in p.parts if re.match(r"word/header\d*\.xml$", n)]
    if not hd: return False
    del p.parts[hd[0]]; return True

@fault("F08", "package", "package", "relationship of unknown type with bad target", "relationship integrity")
def _(p):
    rel = rels(p); r = etree.SubElement(rel, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
    r.set("Id", "rId7777"); r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"); r.set("Target", "media/nothere.png"); return True

@fault("F09", "package", "wellformed", "document.xml truncated (unclosed tag)", "XML well-formedness")
def _(p):
    p.commit(); d = p.parts["word/document.xml"]; p.parts["word/document.xml"] = d[: len(d) - 40]; p._trees.pop("word/document.xml", None); return True

@fault("F10", "package", "package", "document part has a wrong content type", "content type of main part")
def _(p):
    ct = p.tree("[Content_Types].xml")
    for o in ct:
        if o.get("PartName") == "/word/document.xml":
            o.set("ContentType", "application/xml"); return True
    return False


# ---- valid controls (must NOT be flagged) ---------------------------------------
CONTROLS = []
def control(cid, title):
    def deco(fn):
        CONTROLS.append(dict(id=cid, title=title, fn=fn)); return fn
    return deco

def text_nodes(p): return p.x("//w:t")

@control("V01", "reword the text of one run")
def _(p): t = text_nodes(p)[0]; t.text = (t.text or "") + " edited"; return True
@control("V02", "add valid bold to a run")
def _(p): return ensure_rpr_child(p, "b") is not None
@control("V03", "add valid italic to a run")
def _(p): return ensure_rpr_child(p, "i") is not None
@control("V04", "add valid color to a run")
def _(p): return ensure_rpr_child(p, "color", val="1F4E79") is not None
@control("V05", "add valid size to a run")
def _(p): return ensure_rpr_child(p, "sz", val="28") is not None
@control("V06", "add valid highlight")
def _(p): return ensure_rpr_child(p, "highlight", val="yellow") is not None
@control("V07", "valid jc on paragraph")
def _(p):
    ppr = first(p, "//w:pPr[not(w:jc)]")
    if ppr is None: return False
    # jc must come after spacing/ind in sequence: append is fine when no later siblings exist; insert before rPr if present
    rpr = ppr.find(q("w:rPr")); j = E("jc", val="center")
    if rpr is not None: move_before(j, rpr)
    else: ppr.append(j)
    return True
@control("V08", "add a valid empty paragraph at the end of body before sectPr")
def _(p):
    b = first(p, "//w:body"); s = b.find(q("w:sectPr")); para = E("p")
    if s is not None: move_before(para, s)
    else: b.append(para)
    return True
@control("V09", "add a valid hyperlink-free run")
def _(p):
    para = first(p, "//w:p"); r = E("r"); t = E("t"); t.text = " added"; r.append(t); para.append(r); return True
@control("V10", "add a valid bookmark pair")
def _(p):
    para = first(p, "//w:p"); para.append(E("bookmarkStart", id="555", name="ok_bm")); para.append(E("bookmarkEnd", id="555")); return True
@control("V11", "valid spacing")
def _(p):
    ppr = first(p, "//w:pPr[not(w:spacing)]")
    if ppr is None: return False
    sp = E("spacing", after="120")
    # spacing goes before ind/jc/rPr if present
    for tag in ("ind", "contextualSpacing", "jc", "rPr", "numPr"):
        nxt = ppr.find(q("w:" + tag))
        if nxt is not None and tag != "numPr":
            move_before(sp, nxt); return True
    ppr.append(sp); return True
@control("V12", "valid underline single")
def _(p): return ensure_rpr_child(p, "u", val="single") is not None
@control("V13", "valid strike")
def _(p): return ensure_rpr_child(p, "strike") is not None
@control("V14", "reword a table cell")
def _(p):
    t = first(p, "//w:tc//w:t")
    if t is None: return False
    t.text = (t.text or "") + " x"; return True
@control("V15", "duplicate a paragraph (same content)")
def _(p):
    para = first(p, "//w:body/w:p[w:r]")
    if para is None: return False
    para.addnext(copy.deepcopy(para)); return True
@control("V16", "delete one body paragraph")
def _(p):
    ps = p.x("//w:body/w:p[w:r]")
    if len(ps) < 3: return False
    ps[1].getparent().remove(ps[1]); return True
@control("V17", "valid small caps")
def _(p): return ensure_rpr_child(p, "smallCaps") is not None
@control("V18", "valid vertAlign superscript")
def _(p): return ensure_rpr_child(p, "vertAlign", val="superscript") is not None
@control("V19", "valid tracked insertion with all required attributes")
def _(p):
    para = first(p, "//w:p"); ins = E("ins", id="9001", author="Bench", date="2026-01-01T00:00:00Z"); r = E("r"); t = E("t"); t.text = "x"; r.append(t); ins.append(r); para.append(ins); return True
@control("V20", "valid page margin tweak")
def _(p):
    m = first(p, "//w:pgMar")
    if m is None: return False
    m.set(q("w:left"), "1500"); return True


def main():
    corpus, out = sys.argv[1], sys.argv[2]
    os.makedirs(out, exist_ok=True)
    manifest = {"faults": [], "controls": [], "bases": BASES}
    for base in BASES:
        src = os.path.join(corpus, base + ".docx")
        for f in FAULTS:
            pkg = Pkg(src)
            try:
                ok = f["fn"](pkg)
            except Exception as e:
                ok = False
            entry = {"id": f["id"], "category": f["category"], "kind": f["kind"], "title": f["title"], "rule": f["rule"], "base": base, "applied": bool(ok)}
            if ok:
                name = f"{base}__{f['id']}.docx"
                pkg.save(os.path.join(out, name)); entry["file"] = name
            manifest["faults"].append(entry)
        for c in CONTROLS:
            pkg = Pkg(src)
            try: ok = c["fn"](pkg)
            except Exception: ok = False
            entry = {"id": c["id"], "title": c["title"], "base": base, "applied": bool(ok)}
            if ok:
                name = f"{base}__{c['id']}.docx"; pkg.save(os.path.join(out, name)); entry["file"] = name
            manifest["controls"].append(entry)
    json.dump(manifest, open(os.path.join(out, "faults.json"), "w"), indent=1)
    ap = sum(1 for f in manifest["faults"] if f["applied"]); ac = sum(1 for c in manifest["controls"] if c["applied"])
    print(f"faults applied {ap}/{len(manifest['faults'])}; controls applied {ac}/{len(manifest['controls'])}")
    for f in manifest["faults"]:
        if not f["applied"]: print("  not applicable:", f["base"], f["id"])

if __name__ == "__main__":
    main()
