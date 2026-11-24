import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import {defineConfig} from 'vite';
import {VitePWA} from 'vite-plugin-pwa';

const PRODUCTION_CSP = [
  "default-src 'self'",
  "script-src 'self' 'wasm-unsafe-eval' https://cdn.jsdelivr.net",
  "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
  "font-src 'self' https://fonts.gstatic.com data:",
  "img-src 'self' data: blob:",
  "worker-src 'self' blob: https://cdn.jsdelivr.net",
  "connect-src 'self' https://cdn.jsdelivr.net https://tessdata.projectnaptha.com",
  "object-src 'none'",
  "base-uri 'self'",
  "form-action 'self'",
].join('; ');

const SECURITY_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
};

export default defineConfig(() => {
  return {
    plugins: [
      react(),
      tailwindcss(),
      VitePWA({
        registerType: 'prompt',
        manifestFilename: 'manifest.webmanifest',
        includeAssets: [
          'offline.html',
          'apple-touch-icon.png',
          'icons/icon.svg',
          'icons/logo.svg',
          'icons/icon-16.png',
          'icons/icon-32.png',
          'icons/icon-48.png',
          'icons/icon-72.png',
          'icons/icon-96.png',
          'icons/icon-144.png',
          'icons/icon-192.png',
          'icons/icon-512.png',
          'icons/icon-maskable-512.png',
          'screenshots/wide.png',
          'screenshots/narrow.png',
          'workers/pdf.worker.min.mjs',
          'wasm/qpdf.wasm',
          'fonts/*.woff2',
          'fonts/*.ttf',
        ],
        manifest: {
          id: '/?source=pwa',
          name: 'eSZ OFFICE TOOLS',
          short_name: 'eSZ Tools',
          description:
            'Free, offline-first Progressive Web App office suite for PDFs, Documents, Spreadsheets, and Presentations.',
          start_url: '/?source=pwa',
          scope: '/',
          display: 'standalone',
          orientation: 'any',
          theme_color: '#08060E',
          background_color: '#08060E',
          lang: 'en',
          categories: ['productivity', 'utilities'],
          icons: [
            {src: '/icons/icon-16.png', sizes: '16x16', type: 'image/png', purpose: 'any'},
            {src: '/icons/icon-32.png', sizes: '32x32', type: 'image/png', purpose: 'any'},
            {src: '/icons/icon-48.png', sizes: '48x48', type: 'image/png', purpose: 'any'},
            {src: '/icons/icon-72.png', sizes: '72x72', type: 'image/png', purpose: 'any'},
            {src: '/icons/icon-96.png', sizes: '96x96', type: 'image/png', purpose: 'any'},
            {src: '/icons/icon-144.png', sizes: '144x144', type: 'image/png', purpose: 'any'},
            {src: '/icons/icon-192.png', sizes: '192x192', type: 'image/png', purpose: 'any'},
            {src: '/icons/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'any'},
            {
              src: '/icons/icon-maskable-512.png',
              sizes: '512x512',
              type: 'image/png',
              purpose: 'maskable',
            },
          ],
          screenshots: [
            {
              src: '/screenshots/wide.png',
              sizes: '1280x720',
              type: 'image/png',
              form_factor: 'wide',
              label: 'eSZ OFFICE Desktop Workspace and Local File Manager',
            },
            {
              src: '/screenshots/narrow.png',
              sizes: '720x1280',
              type: 'image/png',
              form_factor: 'narrow',
              label: 'eSZ OFFICE Mobile Touch Workspace',
            },
          ],
          shortcuts: [
            {
              name: 'New PDF',
              short_name: 'New PDF',
              description: 'Create a new PDF document',
              url: '/?action=new&type=pdf',
              icons: [{src: '/icons/icon-96.png', sizes: '96x96', type: 'image/png'}],
            },
            {
              name: 'New Doc',
              short_name: 'New Doc',
              description: 'Create a new text document',
              url: '/?action=new&type=docs',
              icons: [{src: '/icons/icon-96.png', sizes: '96x96', type: 'image/png'}],
            },
            {
              name: 'New Sheet',
              short_name: 'New Sheet',
              description: 'Create a new spreadsheet',
              url: '/?action=new&type=sheets',
              icons: [{src: '/icons/icon-96.png', sizes: '96x96', type: 'image/png'}],
            },
            {
              name: 'New Slides',
              short_name: 'New Slides',
              description: 'Create a new presentation',
              url: '/?action=new&type=slides',
              icons: [{src: '/icons/icon-96.png', sizes: '96x96', type: 'image/png'}],
            },
          ],
          file_handlers: [
            {
              action: '/workspace',
              accept: {
                'application/pdf': ['.pdf'],
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
                'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                'text/csv': ['.csv'],
              },
            },
          ],
          share_target: {
            action: '/?share=1',
            method: 'POST',
            enctype: 'multipart/form-data',
            params: {
              title: 'title',
              text: 'text',
              files: [
                {
                  name: 'files',
                  accept: [
                    '.pdf',
                    '.docx',
                    '.pptx',
                    '.xlsx',
                    '.csv',
                    'application/pdf',
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'text/csv',
                  ],
                },
              ],
            },
          },
        },
        workbox: {
          maximumFileSizeToCacheInBytes: 5 * 1024 * 1024,
          globPatterns: ['**/*.{js,css,html,ico,png,svg,woff,woff2,ttf,wasm,mjs}'],
          navigateFallback: '/index.html',
          navigateFallbackDenylist: [/^\/api/, /^\/\.well-known/],
          cleanupOutdatedCaches: true,
          runtimeCaching: [
            {
              urlPattern: /^https:\/\/fonts\.googleapis\.com\/.*/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'google-fonts-stylesheets',
                expiration: {
                  maxEntries: 10,
                  maxAgeSeconds: 60 * 60 * 24 * 365,
                },
                cacheableResponse: {statuses: [0, 200]},
              },
            },
            {
              urlPattern: /^https:\/\/fonts\.gstatic\.com\/.*/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'google-fonts-webfonts',
                expiration: {
                  maxEntries: 20,
                  maxAgeSeconds: 60 * 60 * 24 * 365,
                },
                cacheableResponse: {statuses: [0, 200]},
              },
            },
            {
              urlPattern: /\.(?:wasm|worker\.js|mjs)$/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'office-heavy-workers-wasm',
                expiration: {
                  maxEntries: 30,
                  maxAgeSeconds: 60 * 60 * 24 * 90,
                },
                cacheableResponse: {statuses: [0, 200]},
              },
            },
            {
              urlPattern: /^https:\/\/(?:cdn\.jsdelivr\.net|tessdata\.projectnaptha\.com)\/.*/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'ocr-traineddata-cache',
                expiration: {
                  maxEntries: 20,
                  maxAgeSeconds: 60 * 60 * 24 * 180,
                },
                cacheableResponse: {statuses: [0, 200]},
              },
            },
          ],
        },
        devOptions: {
          enabled: true,
          type: 'module',
        },
      }),
      {
        name: 'inject-production-csp',
        apply: 'build',
        transformIndexHtml(html) {
          return html.replace(
            '<meta charset="UTF-8" />',
            `<meta charset="UTF-8" />\n    <meta http-equiv="Content-Security-Policy" content="${PRODUCTION_CSP}" />`,
          );
        },
      },
      // Copies src/assets into dist/ so root-relative URLs like
      // "/src/assets/images/..." (e.g. generate_image output) still resolve
      // in production builds. Do not remove.
      {
        name: 'copy-unbundled-src-assets',
        apply: 'build',
        closeBundle(error?: Error): void {
          if (error) {
            return;
          }
          const distDir = path.resolve(__dirname, 'dist');
          const srcAssets = path.resolve(__dirname, 'src/assets');
          const distAssets = path.resolve(distDir, 'src/assets');
          if (fs.existsSync(distDir) && fs.existsSync(srcAssets)) {
            fs.cpSync(srcAssets, distAssets, {recursive: true, force: true});
          }
        },
      },
    ],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
      dedupe: [
        '@tiptap/core',
        '@tiptap/pm',
        '@tiptap/react',
        'prosemirror-state',
        'prosemirror-view',
        'prosemirror-model',
        'prosemirror-transform',
      ],
    },
    optimizeDeps: {
      exclude: ['@mssio/qpdf-wasm'],
      include: [
        '@tiptap/core',
        '@tiptap/react',
        '@tiptap/starter-kit',
        '@tiptap/pm/state',
        '@tiptap/pm/view',
        '@tiptap/pm/model',
      ],
    },
    server: {
      headers: SECURITY_HEADERS,
      hmr: process.env.DISABLE_HMR !== 'true',
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
    preview: {
      headers: {
        ...SECURITY_HEADERS,
        'Content-Security-Policy': PRODUCTION_CSP,
      },
    },
  };
});
