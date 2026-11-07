import { readFileSync } from 'node:fs';
import { AdmZip, DocxParser, ToonGenerator } from '../packages/engine/src/index';

async function main(): Promise<void> {
  const file = process.argv[2];
  if (!file) throw new Error('Usage: npx tsx scripts/measure-toon.ts path/to/document.docx');
  const bytes = readFileSync(file);
  const zip = new AdmZip(bytes);
  const xml = zip.readAsText('word/document.xml');
  const parser = new DocxParser(bytes);
  const ast = await parser.getAST();
  const generator = new ToonGenerator();
  const outline = generator.renderToonText(generator.toon(ast, {fidelity:'standard'}));
  const estimate = (text: string): number => Math.round(text.length / 3.8);
  process.stdout.write(`document.xml: ${xml.length} characters, about ${estimate(xml)} tokens\n`);
  process.stdout.write(`TOON: ${outline.length} characters, about ${estimate(outline)} tokens\n`);
}
main().catch((error: unknown) => { process.stderr.write(`${String(error)}\n`); process.exitCode = 1; });
