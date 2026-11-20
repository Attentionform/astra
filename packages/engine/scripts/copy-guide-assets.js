// eslint-disable-next-line @typescript-eslint/no-require-imports -- intentional: CJS build script (package is CommonJS); require is correct here
const fs = require('fs');
// eslint-disable-next-line @typescript-eslint/no-require-imports -- intentional: CJS build script (package is CommonJS); require is correct here
const path = require('path');

function copyGuideAssets() {
  try {
    const srcGuides = path.resolve(__dirname, '..', 'templates', 'guides');
    const targetGuides = path.resolve(__dirname, '..', 'dist', 'templates', 'guides');

    if (!fs.existsSync(srcGuides)) {
      return;
    }

    fs.mkdirSync(targetGuides, { recursive: true });
    for (const entry of fs.readdirSync(srcGuides)) {
      if (!entry.toLowerCase().endsWith('.md')) continue;
      fs.copyFileSync(path.join(srcGuides, entry), path.join(targetGuides, entry));
    }
  } catch {
    /* build-time copy stays best-effort; missing guides surface as resolve failures */
  }
}

function copyTemplateFiles() {
  try {
    const src = path.resolve(__dirname, '..', 'templates');
    const target = path.resolve(__dirname, '..', 'dist', 'templates');
    if (!fs.existsSync(src)) return;
    fs.mkdirSync(target, { recursive: true });
    for (const entry of fs.readdirSync(src)) {
      if (entry.toLowerCase().endsWith('.docx')) fs.copyFileSync(path.join(src, entry), path.join(target, entry));
    }
  } catch (err) {
    // A missing template file is an error at run time (createNewDocument names the file); report it here too.
    console.error('[copy-templates] failed:', err && err.message);
    process.exitCode = 1;
  }
}

copyGuideAssets();
copyTemplateFiles();
