import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  build: {
    lib: { entry: 'src/main.jsx', name: 'PainelDemandas', formats: ['iife'], fileName: () => 'dashboard.js' },
    cssCodeSplit: false,
    rollupOptions: { output: { assetFileNames: asset => asset.name?.endsWith('.css') ? 'dashboard.css' : 'assets/[name]-[hash][extname]' } },
  },
});
