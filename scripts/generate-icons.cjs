const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

// Ensure public directory exists
const publicDir = path.resolve(__dirname, '../public');
if (!fs.existsSync(publicDir)) {
  fs.mkdirSync(publicDir, { recursive: true });
}

// 1. Create public/icon.svg
const iconSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="512" height="512">
  <defs>
    <linearGradient id="bgGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0f172a" />
      <stop offset="50%" stop-color="#1e293b" />
      <stop offset="100%" stop-color="#0284c7" />
    </linearGradient>
    <linearGradient id="boltGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#38bdf8" />
      <stop offset="50%" stop-color="#3b82f6" />
      <stop offset="100%" stop-color="#10b981" />
    </linearGradient>
    <linearGradient id="shieldGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#38bdf8" stop-opacity="0.3" />
      <stop offset="100%" stop-color="#10b981" stop-opacity="0.1" />
    </linearGradient>
    <filter id="glow" x="-20%" y="-20%" width="140%" height="140%">
      <feGaussianBlur stdDeviation="12" result="blur" />
      <feComposite in="SourceGraphic" in2="blur" operator="over" />
    </filter>
  </defs>

  <!-- Background rounded rect -->
  <rect width="512" height="512" rx="112" fill="url(#bgGrad)" />

  <!-- Outer Energy Ring -->
  <circle cx="256" cy="256" r="190" fill="none" stroke="url(#boltGrad)" stroke-width="8" stroke-dasharray="16 12" opacity="0.6" />
  <circle cx="256" cy="256" r="160" fill="url(#shieldGrad)" stroke="#38bdf8" stroke-width="3" opacity="0.8" />

  <!-- Lightning Bolt Energy Symbol -->
  <path d="M280 80 L160 270 L250 270 L230 432 L352 242 L262 242 Z"
        fill="url(#boltGrad)"
        filter="url(#glow)"
        stroke="#ffffff"
        stroke-width="4"
        stroke-linejoin="round" />

  <!-- Bottom Brand Label -->
  <rect x="146" y="420" width="220" height="48" rx="24" fill="#0f172a" stroke="#38bdf8" stroke-width="2" />
  <text x="256" y="452" font-family="system-ui, -apple-system, sans-serif" font-size="24" font-weight="900" fill="#ffffff" text-anchor="middle" letter-spacing="4">TEV CMMS</text>
</svg>`;

fs.writeFileSync(path.join(publicDir, 'icon.svg'), iconSvg, 'utf8');
console.log('Created public/icon.svg');

// PNG Generator Helper
function generatePng(size, isMaskable = false) {
  const width = size;
  const height = size;

  function crc32(buf) {
    let table = [];
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) {
        if (c & 1) c = 0xedb88320 ^ (c >>> 1);
        else c = c >>> 1;
      }
      table[n] = c;
    }
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i++) {
      c = table[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    }
    return (c ^ 0xffffffff) >>> 0;
  }

  function chunk(type, data) {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length, 0);
    const typeAndData = Buffer.concat([Buffer.from(type), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(typeAndData), 0);
    return Buffer.concat([len, typeAndData, crc]);
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6; // RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const rowSize = 1 + width * 4;
  const rawData = Buffer.alloc(rowSize * height);

  const cx = width / 2;
  const cy = height / 2;
  const radius = width / 2;
  const cornerRadius = isMaskable ? 0 : width * 0.22;

  for (let y = 0; y < height; y++) {
    const rowOffset = y * rowSize;
    rawData[rowOffset] = 0;

    for (let x = 0; x < width; x++) {
      const pxOffset = rowOffset + 1 + x * 4;

      // Check rounded corner bounding box if not maskable
      let inBounds = true;
      if (!isMaskable) {
        let dx = Math.max(Math.abs(x - cx) - (width / 2 - cornerRadius), 0);
        let dy = Math.max(Math.abs(y - cy) - (height / 2 - cornerRadius), 0);
        if (dx * dx + dy * dy > cornerRadius * cornerRadius) {
          inBounds = false;
        }
      }

      if (!inBounds) {
        rawData[pxOffset] = 0;
        rawData[pxOffset + 1] = 0;
        rawData[pxOffset + 2] = 0;
        rawData[pxOffset + 3] = 0;
        continue;
      }

      // Background Gradient (#0f172a to #0284c7)
      const gradT = (x + y) / (width + height);
      let r = Math.round(15 + gradT * (2 - 15));
      let g = Math.round(23 + gradT * (132 - 23));
      let b = Math.round(42 + gradT * (199 - 42));

      // Circular Ring
      const distFromCenter = Math.sqrt((x - cx) * (x - cx) + (y - cy) * (y - cy));
      const ringR = width * (isMaskable ? 0.32 : 0.36);
      if (Math.abs(distFromCenter - ringR) < width * 0.015) {
        r = 56; g = 189; b = 248; // Cyan ring
      } else if (distFromCenter < ringR) {
        // Inner tint
        r = Math.min(255, r + 15);
        g = Math.min(255, g + 25);
        b = Math.min(255, b + 35);
      }

      // Simplified Lightning Bolt Shape in center
      // Points scaled to size:
      // (280, 80) -> (160, 270) -> (250, 270) -> (230, 432) -> (352, 242) -> (262, 242)
      const scale = (isMaskable ? 0.65 : 0.75) * (width / 512);
      const nx = (x - cx) / scale + 256;
      const ny = (y - cy) / scale + 256;

      // Approximate lightning polygon
      let inBolt = false;
      // Top triangle: 280,80 -> 160,270 -> 262,242
      // Middle strip
      // Bottom triangle: 250,270 -> 230,432 -> 352,242
      if (ny >= 100 && ny <= 410) {
        const topEdge = 100 + (nx - 160) * 1.5;
        const bottomEdge = 420 - (350 - nx) * 1.6;
        if (nx >= 170 && nx <= 340) {
          if (ny <= 270 && nx >= 160 + (ny - 100) * 0.4 && nx <= 300) {
            inBolt = true;
          }
          if (ny >= 240 && nx >= 210 && nx <= 250 + (410 - ny) * 0.6) {
            inBolt = true;
          }
        }
      }

      if (inBolt) {
        const boltGrad = (ny - 100) / 310;
        r = Math.round(56 + boltGrad * (16 - 56));
        g = Math.round(189 + boltGrad * (185 - 189));
        b = Math.round(248 + boltGrad * (129 - 248));
      }

      rawData[pxOffset] = Math.max(0, Math.min(255, r));
      rawData[pxOffset + 1] = Math.max(0, Math.min(255, g));
      rawData[pxOffset + 2] = Math.max(0, Math.min(255, b));
      rawData[pxOffset + 3] = 255;
    }
  }

  const compressed = zlib.deflateSync(rawData);
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  return Buffer.concat([signature, chunk('IHDR', ihdr), chunk('IDAT', compressed), chunk('IEND', Buffer.alloc(0))]);
}

// Generate all required PWA icons
fs.writeFileSync(path.join(publicDir, 'apple-touch-icon.png'), generatePng(180, false));
console.log('Created public/apple-touch-icon.png (180x180)');

fs.writeFileSync(path.join(publicDir, 'pwa-192x192.png'), generatePng(192, false));
console.log('Created public/pwa-192x192.png (192x192)');

fs.writeFileSync(path.join(publicDir, 'pwa-512x512.png'), generatePng(512, false));
console.log('Created public/pwa-512x512.png (512x512)');

fs.writeFileSync(path.join(publicDir, 'pwa-maskable-512x512.png'), generatePng(512, true));
console.log('Created public/pwa-maskable-512x512.png (512x512 maskable)');

// Create favicon.ico from 32x32 PNG
fs.writeFileSync(path.join(publicDir, 'favicon.ico'), generatePng(32, false));
console.log('Created public/favicon.ico');
