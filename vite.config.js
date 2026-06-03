import { defineConfig } from 'vite';
import { viteSingleFile } from 'vite-plugin-singlefile';
import path from 'path';

export default defineConfig({
  root: path.resolve(__dirname, 'src/web'),
  plugins: [viteSingleFile()],
  build: {
    outDir: path.resolve(__dirname, 'dist'),
    emptyOutDir: true,
    assetsInlineLimit: 100000000, // Make sure all assets are inlined
    chunkSizeWarningLimit: 10000,
    cssCodeSplit: false
  }
});
