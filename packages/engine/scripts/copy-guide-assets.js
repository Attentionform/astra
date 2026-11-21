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

copyGuideAssets();
