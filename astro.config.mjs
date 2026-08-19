import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

// Local by default. To publish on GitHub Pages later, set these two env vars —
// nothing else in the codebase needs to change, because every internal link is
// built from import.meta.env.BASE_URL (see src/lib/url.js).
//
//   Project site   SITE_URL=https://<user>.github.io  SITE_BASE=/engineering-portfolio
//   Custom domain  SITE_URL=https://example.com       SITE_BASE=/
//
// publish-site.bat sets them for you.
const site = process.env.SITE_URL || 'http://localhost:4321';
const base = process.env.SITE_BASE || '/';

export default defineConfig({
  site,
  base,
  integrations: [sitemap()],
  build: { format: 'directory' },
  devToolbar: { enabled: false },
});
