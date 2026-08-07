import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  // Use relative asset URLs so the generated site also works when Cloudflare
  // Pages serves it below a path (for example, on preview deployments).
  base: './',
  plugins: [react()],
});
