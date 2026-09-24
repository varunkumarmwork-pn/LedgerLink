import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    open: true,
    // The Python backend answers everything under /api. Port 8001 because 8000 is
    // already used by another app on this machine.
    proxy: { '/api': 'http://127.0.0.1:8001' },
  },
});
