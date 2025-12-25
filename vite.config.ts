import { defineConfig } from 'vite';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

// Office Add-in dev certs location
const certPath = path.join(os.homedir(), '.office-addin-dev-certs');

// Try to load existing certificates
function getHttpsConfig() {
  try {
    const keyPath = path.join(certPath, 'localhost.key');
    const certFilePath = path.join(certPath, 'localhost.crt');
    
    if (fs.existsSync(keyPath) && fs.existsSync(certFilePath)) {
      return {
        key: fs.readFileSync(keyPath),
        cert: fs.readFileSync(certFilePath),
      };
    }
  } catch (e) {
    console.warn('Could not load SSL certificates. Run: npx office-addin-dev-certs install');
  }
  return undefined;
}

export default defineConfig({
  server: {
    port: 3000,
    https: getHttpsConfig(),
    cors: true,
  },
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        main: 'index.html',
      },
    },
  },
});
