import fs from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
import { load } from 'cheerio';
import xlsx from 'xlsx';
import { chromium } from 'playwright';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const args = process.argv.slice(2);
const singlePage = args.includes('--single-page');
const remainingArgs = args.filter(arg => arg !== '--single-page');

if (remainingArgs.length < 1) {
  console.error('Usage: node extract-images.js [--single-page] <url>');
  process.exit(1);
}

const startUrl = remainingArgs[0];

function normalizeUrl(urlString, baseUrl) {
  try {
    const resolved = new URL(urlString, baseUrl);
    resolved.hash = '';
    if (resolved.pathname.endsWith('/') && resolved.pathname !== '/') {
      resolved.pathname = resolved.pathname.replace(/\/+$/, '');
    }
    return resolved.toString();
  } catch {
    return null;
  }
}

function isSameOrigin(urlA, urlB) {
  try {
    const a = new URL(urlA);
    const b = new URL(urlB);
    return a.protocol === b.protocol && a.hostname === b.hostname && a.port === b.port;
  } catch {
    return false;
  }
}

async function fetchHtml(page, url) {
  const response = await page.goto(url, {
    waitUntil: 'domcontentloaded',
    timeout: 30000
  });
  if (!response) {
    throw new Error(`Failed to load ${url}: no response`);
  }
  if (!response.ok()) {
    throw new Error(`Failed to fetch ${url}: ${response.status()} ${response.statusText()}`);
  }

  await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
  return await page.content();
}

function extractLinks(html, pageUrl) {
  const $ = load(html);
  const images = new Set();
  const routes = new Set();

  $('img[src]').each((_, img) => {
    const src = $(img).attr('src');
    const norm = normalizeUrl(src, pageUrl);
    if (norm) {
      images.add(norm);
    }
  });

  $('a[href]').each((_, link) => {
    const href = $(link).attr('href').trim();
    if (!href || href.startsWith('mailto:') || href.startsWith('tel:') || href.startsWith('javascript:') || href.startsWith('#')) {
      return;
    }
    const norm = normalizeUrl(href, pageUrl);
    if (norm && isSameOrigin(norm, pageUrl)) {
      routes.add(norm);
    }
  });

  return {
    images: Array.from(images),
    routes: Array.from(routes)
  };
}

async function crawl(startUrl, singlePage) {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
    locale: 'en-US',
    viewport: { width: 1280, height: 800 },
    extraHTTPHeaders: {
      Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
      'Accept-Language': 'en-US,en;q=0.9',
      'Cache-Control': 'no-cache',
      Pragma: 'no-cache'
    }
  });
  const imageSizes = new Map();
  const page = await context.newPage();

  page.on('response', async (response) => {
    try {
      if (response.request().resourceType() !== 'image') {
        return;
      }
      const rawUrl = response.url();
      const normUrl = normalizeUrl(rawUrl, startUrl);
      if (!normUrl) {
        return;
      }
      const headers = response.headers();
      let size = null;
      if (headers['content-length']) {
        const parsed = Number(headers['content-length']);
        if (!Number.isNaN(parsed)) {
          size = parsed;
        }
      }
      if (size === null) {
        const body = await response.body();
        size = body ? body.byteLength : null;
      }
      if (size !== null) {
        imageSizes.set(normUrl, Math.round((size / 1024) * 100) / 100);
      }
    } catch {
      // ignore failures for image response metadata
    }
  });

  const visited = new Set();
  const queue = [startUrl];
  const allImages = [];

  while (queue.length > 0) {
    const currentUrl = queue.shift();
    if (visited.has(currentUrl)) {
      continue;
    }

    console.log(`Crawling: ${currentUrl}`);
    visited.add(currentUrl);

    try {
      const html = await fetchHtml(page, currentUrl);
      const { routes } = extractLinks(html, currentUrl);

      const pageImages = await page.$$eval('img[src]', (imgs) =>
        imgs.map((img) => ({
          src: img.getAttribute('src'),
          alt: img.getAttribute('alt') || '',
          width: img.naturalWidth || img.clientWidth || 0
        }))
      );

      const seenImages = new Set();
      pageImages.forEach(({ src, alt, width }) => {
        const normalizedImage = normalizeUrl(src, currentUrl);
        if (!normalizedImage || seenImages.has(normalizedImage)) {
          return;
        }
        seenImages.add(normalizedImage);
        allImages.push({
          route: currentUrl,
          image: normalizedImage,
          alt,
          size: imageSizes.get(normalizedImage) ?? null,
          width: width || null
        });
      });

      if (!singlePage) {
        routes.forEach((route) => {
          if (!visited.has(route) && !queue.includes(route)) {
            queue.push(route);
          }
        });
      }
    } catch (error) {
      console.warn(`Warning: could not crawl ${currentUrl} — ${error.message}`);
    }
  }

  await browser.close();
  return allImages;
}

function writeExcel(allImages, hostname) {
  const workbook = xlsx.utils.book_new();

  const sheet = [ ['RouteUrl', 'ImageUrl', 'ImageAltText', 'ImageSizeKB', 'ImageWidth'] ];
  for (const entry of allImages) {
    sheet.push([entry.route, entry.image, entry.alt, entry.size, entry.width]);
  }

  workbook.SheetNames.push('Data');
  workbook.Sheets.Data = xlsx.utils.aoa_to_sheet(sheet);

  const outputPath = path.join(__dirname, `output-${hostname}.xlsx`);
  xlsx.writeFile(workbook, outputPath);
  return outputPath;
}

function getHostname(urlString) {
  try {
    return new URL(urlString).hostname;
  } catch {
    return 'output';
  }
}

(async () => {
  try {
    const normalizedStartUrl = normalizeUrl(startUrl, startUrl);
    if (!normalizedStartUrl) {
      throw new Error('Invalid start URL. Use a full URL like https://example.com/');
    }

    const allImages = await crawl(normalizedStartUrl, singlePage);
    const hostname = getHostname(normalizedStartUrl);
    const outputPath = writeExcel(allImages, hostname);

    console.log(`\nFinished crawling and extracting ${allImages.length} image(s).`);
    console.log(`Excel output saved to: ${outputPath}`);
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }
})();
