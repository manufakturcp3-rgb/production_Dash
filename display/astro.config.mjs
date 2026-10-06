import { defineConfig } from 'astro/config';

// https://astro.build/config
export default defineConfig({
  output: 'static',
  // Build sebagai static site, cocok untuk deploy di Vercel
});
