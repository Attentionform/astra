#!/usr/bin/env python3
"""Control for E1: parse every XML part with lxml, serialise it, write a new zip. No product involved.
CLI: control_reserialize.py in.docx out.docx"""
import sys, zipfile
from lxml import etree
src, dst = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for n in zin.namelist():
        data = zin.read(n)
        if n.endswith((".xml", ".rels")):
            data = etree.tostring(etree.fromstring(data), xml_declaration=True, encoding="UTF-8", standalone=True)
        zout.writestr(n, data)
