import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const source = path.resolve(__dirname, 'dist/index.html');
const dest = path.resolve(__dirname, 'index.html');

try {
  if (fs.existsSync(source)) {
    fs.copyFileSync(source, dest);
    console.log('✨ [SUCCESS] Successfully copied bundled single-file app to the root index.html!');
  } else {
    console.error(`❌ [ERROR] Build output file not found at: ${source}`);
    process.exit(1);
  }
} catch (err) {
  console.error('❌ [ERROR] Error copying build output:', err);
  process.exit(1);
}
