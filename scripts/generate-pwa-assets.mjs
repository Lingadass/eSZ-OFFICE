import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

const outIcons = path.resolve('public/icons');
const outScreens = path.resolve('public/screenshots');
fs.mkdirSync(outIcons, { recursive: true });
fs.mkdirSync(outScreens, { recursive: true });

// CRC32 helper for PNG chunks
const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  crcTable[n] = c >>> 0;
}
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}
function makeChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, 'ascii');
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([len, typeBuf, data, crcBuf]);
}
function encodePNG(width, height, rgba) {
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const stride = width * 4 + 1;
  const raw = Buffer.alloc(stride * height);
  for (let y = 0; y < height; y++) {
    raw[y * stride] = 0; // filter None
    rgba.copy(raw, y * stride + 1, y * width * 4, (y + 1) * width * 4);
  }
  const compressed = zlib.deflateSync(raw, { level: 9 });
  return Buffer.concat([
    signature,
    makeChunk('IHDR', ihdr),
    makeChunk('IDAT', compressed),
    makeChunk('IEND', Buffer.alloc(0)),
  ]);
}

function blendPixel(buf, width, height, x, y, r, g, b, a) {
  if (x < 0 || y < 0 || x >= width || y >= height || a <= 0) return;
  const idx = (y * width + x) * 4;
  const sa = a / 255;
  const da = buf[idx + 3] / 255;
  const outA = sa + da * (1 - sa);
  if (outA <= 0) return;
  buf[idx] = Math.round((r * sa + buf[idx] * da * (1 - sa)) / outA);
  buf[idx + 1] = Math.round((g * sa + buf[idx + 1] * da * (1 - sa)) / outA);
  buf[idx + 2] = Math.round((b * sa + buf[idx + 2] * da * (1 - sa)) / outA);
  buf[idx + 3] = Math.round(outA * 255);
}

function roundedRectDist(px, py, x, y, w, h, rad) {
  const cx = Math.max(x + rad, Math.min(px, x + w - rad));
  const cy = Math.max(y + rad, Math.min(py, y + h - rad));
  return Math.hypot(px - cx, py - cy) - rad;
}

function drawRoundedRect(buf, width, height, x, y, w, h, rad, color) {
  const [r, g, b, a] = color;
  const x0 = Math.max(0, Math.floor(x - 2));
  const y0 = Math.max(0, Math.floor(y - 2));
  const x1 = Math.min(width - 1, Math.ceil(x + w + 2));
  const y1 = Math.min(height - 1, Math.ceil(y + h + 2));
  for (let py = y0; py <= y1; py++) {
    for (let px = x0; px <= x1; px++) {
      const d = roundedRectDist(px + 0.5, py + 0.5, x, y, w, h, rad);
      if (d < 0.5) {
        const cov = Math.max(0, Math.min(1, 0.5 - d));
        blendPixel(buf, width, height, px, py, r, g, b, Math.round(a * cov));
      }
    }
  }
}

function renderIcon(size, maskable = false) {
  const buf = Buffer.alloc(size * size * 4);
  if (maskable) {
    // Full-bleed slate background + inner cobalt square inside 72% safe zone
    for (let i = 0; i < size * size; i++) {
      buf[i * 4] = 15;
      buf[i * 4 + 1] = 23;
      buf[i * 4 + 2] = 42;
      buf[i * 4 + 3] = 255;
    }
    const pad = size * 0.16;
    const inner = size - pad * 2;
    drawRoundedRect(buf, size, size, pad, pad, inner, inner, inner * 0.22, [29, 78, 216, 255]);
    const barX = pad + inner * 0.22;
    const barW = inner * 0.56;
    const barH = Math.max(2, inner * 0.085);
    drawRoundedRect(buf, size, size, barX, pad + inner * 0.27, barW, barH, barH / 2, [255, 255, 255, 255]);
    drawRoundedRect(buf, size, size, barX, pad + inner * 0.46, barW * 0.68, barH, barH / 2, [255, 255, 255, 255]);
    drawRoundedRect(buf, size, size, barX + barW * 0.78, pad + inner * 0.43, barW * 0.22, barH * 1.7, barH * 0.45, [147, 197, 253, 255]);
    drawRoundedRect(buf, size, size, barX, pad + inner * 0.65, barW, barH, barH / 2, [255, 255, 255, 255]);
  } else {
    drawRoundedRect(buf, size, size, 0, 0, size, size, size * 0.22, [15, 23, 42, 255]);
    const inset = size * 0.08;
    const inner = size - inset * 2;
    drawRoundedRect(buf, size, size, inset, inset, inner, inner, inner * 0.2, [29, 78, 216, 255]);
    const barX = size * 0.28;
    const barW = size * 0.44;
    const barH = Math.max(2, size * 0.068);
    drawRoundedRect(buf, size, size, barX, size * 0.31, barW, barH, barH / 2, [255, 255, 255, 255]);
    drawRoundedRect(buf, size, size, barX, size * 0.466, barW * 0.68, barH, barH / 2, [255, 255, 255, 255]);
    drawRoundedRect(buf, size, size, barX + barW * 0.78, size * 0.44, barW * 0.22, barH * 1.75, barH * 0.45, [147, 197, 253, 255]);
    drawRoundedRect(buf, size, size, barX, size * 0.62, barW, barH, barH / 2, [255, 255, 255, 255]);
  }
  return encodePNG(size, size, buf);
}

for (const s of [48, 72, 96, 144, 192, 512]) {
  fs.writeFileSync(path.join(outIcons, `icon-${s}.png`), renderIcon(s, false));
}
fs.writeFileSync(path.join(outIcons, 'icon-maskable-512.png'), renderIcon(512, true));
fs.writeFileSync(path.join('public', 'apple-touch-icon.png'), renderIcon(180, true));

// Seed crisp placeholder screenshot canvases (will be refreshed with live app captures)
function renderScreenshot(w, h) {
  const buf = Buffer.alloc(w * h * 4);
  for (let i = 0; i < w * h; i++) {
    buf[i * 4] = 248;
    buf[i * 4 + 1] = 250;
    buf[i * 4 + 2] = 252;
    buf[i * 4 + 3] = 255;
  }
  // Top bar
  drawRoundedRect(buf, w, h, 0, 0, w, 64, 0, [255, 255, 255, 255]);
  drawRoundedRect(buf, w, h, 24, 16, 32, 32, 8, [29, 78, 216, 255]);
  // Cards
  const cols = w > h ? 4 : 2;
  const pad = 32;
  const gap = 20;
  const cardW = (w - pad * 2 - gap * (cols - 1)) / cols;
  const colors = [
    [220, 38, 38, 255],
    [37, 99, 235, 255],
    [21, 128, 61, 255],
    [217, 119, 6, 255],
  ];
  for (let i = 0; i < cols; i++) {
    const cx = pad + i * (cardW + gap);
    drawRoundedRect(buf, w, h, cx, 96, cardW, 140, 14, [255, 255, 255, 255]);
    drawRoundedRect(buf, w, h, cx + 20, 116, 40, 40, 10, colors[i % 4]);
  }
  return encodePNG(w, h, buf);
}
fs.writeFileSync(path.join(outScreens, 'wide.png'), renderScreenshot(1280, 720));
fs.writeFileSync(path.join(outScreens, 'narrow.png'), renderScreenshot(720, 1280));
console.log('Generated PWA PNG icons and screenshots.');
