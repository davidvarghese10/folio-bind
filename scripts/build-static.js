import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const distDir = path.join(rootDir, 'dist');
const docsDir = path.join(rootDir, 'docs');

// 1. Remove old docs directory and recreate cleanly
if (fs.existsSync(docsDir)) {
  fs.rmSync(docsDir, { recursive: true, force: true });
}
fs.mkdirSync(docsDir, { recursive: true });

// 2. Copy all compiled files from dist to docs
fs.cpSync(distDir, docsDir, { recursive: true });

// 3. Ensure docs/index.html is clean (strip any redirect scripts intended only for root)
const docsIndexPath = path.join(docsDir, 'index.html');
if (fs.existsSync(docsIndexPath)) {
  let docsHtml = fs.readFileSync(docsIndexPath, 'utf-8');
  // Remove redirect scripts and fallback handlers
  docsHtml = docsHtml.replace(/<script>[\s\S]*?\/\/ When served as static[\s\S]*?<\/script>/gi, '');
  docsHtml = docsHtml.replace(/<script>[\s\S]*?\/\/ If served directly[\s\S]*?<\/script>/gi, '');
  docsHtml = docsHtml.replace(/<script>[\s\S]*?\/\/ Fallback message[\s\S]*?<\/script>/gi, '');
  fs.writeFileSync(docsIndexPath, docsHtml, 'utf-8');
}

// 4. Ensure .nojekyll exists in both docs and root to prevent GitHub Pages Jekyll processing
fs.writeFileSync(path.join(docsDir, '.nojekyll'), '', 'utf-8');
fs.writeFileSync(path.join(rootDir, '.nojekyll'), '', 'utf-8');

console.log('✓ Successfully synchronized clean static production build to /docs');
