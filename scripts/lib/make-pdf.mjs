/**
 * Writes small PDFs with chosen page sizes, rotation and crop box, for tests that compare PDF renderers.
 * Each page says "Page N" in big type and has a bar along the bottom, so a render always has ink.
 */
export function makePdf(pages) {
  const objs = [];
  const kids = [];
  objs.push('<</Type/Catalog/Pages 2 0 R>>');
  pages.forEach((_, i) => kids.push(`${3 + i * 2} 0 R`));
  objs.push(`<</Type/Pages/Kids[${kids.join(' ')}]/Count ${pages.length}>>`);
  pages.forEach((p, i) => {
    const extra = (p.rotate ? `/Rotate ${p.rotate}` : '') + (p.crop ? `/CropBox[${p.crop.join(' ')}]` : '');
    objs.push(
      `<</Type/Page/Parent 2 0 R/MediaBox[0 0 ${p.w} ${p.h}]${extra}/Contents ${4 + i * 2} 0 R/Resources<</Font<</F1<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>>>>>>>`
    );
    const content = `BT /F1 40 Tf 30 ${p.h - 80} Td (Page ${i + 1}) Tj ET 0 0 0 rg 20 20 ${p.w - 40} 10 re f`;
    objs.push(`<</Length ${content.length}>>\nstream\n${content}\nendstream`);
  });
  let out = '%PDF-1.4\n';
  const offsets = [];
  objs.forEach((o, i) => {
    offsets.push(out.length);
    out += `${i + 1} 0 obj\n${o}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n`;
  for (const o of offsets) out += `${String(o).padStart(10, '0')} 00000 n \n`;
  return Buffer.from(`${out}trailer\n<</Size ${objs.length + 1}/Root 1 0 R>>\nstartxref\n${xref}\n%%EOF\n`, 'latin1');
}
