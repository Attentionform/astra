import fs from 'fs';
import path from 'path';
import { validateDocxFile } from '@attentionform/astra-engine';

async function main() {
  const galleryDir = path.resolve(process.cwd(), 'test-gallery');
  if (!fs.existsSync(galleryDir)) {
    console.error(`[validate_docs] Error: Gallery directory not found: ${galleryDir}`);
    process.exit(1);
  }

  const files = fs.readdirSync(galleryDir).filter(f => f.endsWith('.docx'));
  if (files.length === 0) {
    console.error(`[validate_docs] Error: No .docx files found in ${galleryDir}. CI gate requires at least 1 document.`);
    process.exit(1);
  }

  console.log(`[validate_docs] Found ${files.length} gallery document(s) to validate in ${galleryDir}`);
  let hasFailure = false;

  for (const file of files) {
    const filePath = path.join(galleryDir, file);
    try {
      const res = await validateDocxFile(filePath);
      if (res.ok) {
        console.log(`  ✓ ${file} (OK)`);
      } else {
        hasFailure = true;
        console.error(`  ✗ ${file} (FAILED with ${res.errors?.length || 0} errors):`);
        for (const err of res.errors || []) {
          console.error(`    - [${err.errorType || 'SchemaError'}] ${err.description || 'Unknown error'} @ ${err.path || err.xPath || ''}`);
        }
      }
    } catch (err: any) {
      hasFailure = true;
      console.error(`  ✗ ${file} (EXCEPTION during validation): ${err?.message || err}`);
    }
  }

  if (hasFailure) {
    console.error(`[validate_docs] Validation FAILED for one or more documents.`);
    process.exit(1);
  }

  console.log(`[validate_docs] All ${files.length} document(s) verified OK: true.`);
  process.exit(0);
}

main().catch(err => {
  console.error(`[validate_docs] Fatal error:`, err);
  process.exit(1);
});
