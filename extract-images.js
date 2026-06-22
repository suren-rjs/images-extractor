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

/**
 * Normalizes a URL, resolving relative URLs against a base URL.
 * Strips hashes and handles trailing slashes for standard URLs.
 * Leaves data URLs untouched. Identifies and filters out local anchor/fragment
 * identifiers and SVG filter/gradient references (including percent-encoded hashes).
 * 
 * @param {string} urlString - The URL string to normalize.
 * @param {string} baseUrl - The base URL to resolve against.
 * @returns {string|null} The normalized URL or null if invalid or a local fragment.
 */
function normalizeUrl(urlString, baseUrl) {
  if (!urlString) return null;
  const trimmed = urlString.trim();
  if (trimmed.startsWith('data:')) {
    return trimmed;
  }
  
  // Replace case-insensitive %23 with # to decode URL-encoded hash fragments
  const decodedUrlStr = trimmed.replace(/%23/gi, '#');
  if (decodedUrlStr.startsWith('#')) {
    return null;
  }

  try {
    const resolved = new URL(decodedUrlStr, baseUrl);

    // If it resolves to the page's own URL but with a hash, it's a local anchor/filter reference
    if (baseUrl) {
      try {
        const baseObj = new URL(baseUrl);
        if (resolved.origin === baseObj.origin && resolved.pathname === baseObj.pathname && resolved.hash !== '') {
          return null;
        }
      } catch {
        // ignore base URL parsing errors
      }
    }

    resolved.hash = '';
    if (resolved.pathname.endsWith('/') && resolved.pathname !== '/') {
      resolved.pathname = resolved.pathname.replace(/\/+$/, '');
    }
    return resolved.toString();
  } catch {
    return null;
  }
}

/**
 * Checks if two URLs have the same origin (protocol, hostname, and port).
 * 
 * @param {string} urlA - The first URL.
 * @param {string} urlB - The second URL.
 * @returns {boolean} True if they have the same origin.
 */
function isSameOrigin(urlA, urlB) {
  try {
    const a = new URL(urlA);
    const b = new URL(urlB);
    return a.protocol === b.protocol && a.hostname === b.hostname && a.port === b.port;
  } catch {
    return false;
  }
}

/**
 * Navigates a browser page to a URL and waits for it to load.
 * 
 * @param {import('playwright').Page} page - The Playwright page object.
 * @param {string} url - The URL to navigate to.
 * @returns {Promise<string>} The HTML content of the page.
 */
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

/**
 * Scroll the page to the bottom to trigger lazy loading of images.
 * 
 * @param {import('playwright').Page} page - The Playwright page.
 * @returns {Promise<void>}
 */
async function autoScroll(page) {
  await page.evaluate(async () => {
    await new Promise((resolve) => {
      let totalHeight = 0;
      const distance = 150;
      const maxScrollHeight = 15000; // Limit to prevent infinite scrolling
      const timer = setInterval(() => {
        const scrollHeight = document.body.scrollHeight;
        window.scrollBy(0, distance);
        totalHeight += distance;

        if (totalHeight >= scrollHeight || totalHeight > maxScrollHeight) {
          clearInterval(timer);
          resolve();
        }
      }, 40);
    });
  });
  // Wait brief moment for layout/lazy resources to fetch
  await page.waitForTimeout(1500);
}

/**
 * Extracts links (routes) from HTML for crawl discovery.
 * 
 * @param {string} html - The HTML content.
 * @param {string} pageUrl - The current page URL.
 * @returns {{routes: string[]}} List of same-origin subpages to crawl.
 */
function extractRoutes(html, pageUrl) {
  const $ = load(html);
  const routes = new Set();

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
    routes: Array.from(routes)
  };
}

/**
 * Browser-side function to locate and extract all potential images.
 * Extracts img, srcset, inline SVGs, CSS backgrounds, link icons, meta images, objects, embeds.
 * 
 * @returns {{src: string, alt: string, width: number|null}[]} List of extracted raw images.
 */
function extractAllImagesFromPage() {
  const images = [];

  function addImage(src, alt = '', width = null) {
    if (!src) return;
    images.push({ src, alt, width });
  }

  const allElements = document.querySelectorAll('*');
  for (const el of allElements) {
    const tagName = el.tagName.toLowerCase();

    // 1. Image tags
    if (tagName === 'img') {
      const src = el.getAttribute('src');
      const alt = el.getAttribute('alt') || '';
      const naturalWidth = el.naturalWidth || el.clientWidth || null;
      addImage(src, alt, naturalWidth);

      // Check srcset attribute
      const srcset = el.getAttribute('srcset');
      if (srcset) {
        srcset.split(',').forEach(part => {
          const url = part.trim().split(/\s+/)[0];
          addImage(url, alt, null);
        });
      }

      // Check lazy-loading attributes (e.g., data-src, data-lazy, original-src, etc.)
      for (const attr of el.attributes) {
        const name = attr.name.toLowerCase();
        if ((name.startsWith('data-') || name.includes('lazy') || name.includes('src')) &&
            name !== 'src' && name !== 'srcset' && name !== 'srclang') {
          const val = attr.value.trim();
          if (val && (val.startsWith('http') || val.startsWith('/') || val.startsWith('data:') || val.match(/\.(png|jpg|jpeg|gif|svg|webp|avif|bmp|ico)(\?.*)?$/i))) {
            addImage(val, alt, null);
          }
        }
      }
    }

    // 2. Picture source elements
    if (tagName === 'source') {
      const srcset = el.getAttribute('srcset');
      if (srcset) {
        srcset.split(',').forEach(part => {
          const url = part.trim().split(/\s+/)[0];
          addImage(url, '', null);
        });
      }
    }

    // 3. Inline SVG elements (serialized to data URLs)
    if (tagName === 'svg') {
      try {
        const svgString = new XMLSerializer().serializeToString(el);
        const encoded = encodeURIComponent(svgString)
          .replace(/'/g, '%27')
          .replace(/"/g, '%22');
        const dataUrl = `data:image/svg+xml;utf8,${encoded}`;
        const alt = el.getAttribute('aria-label') || el.querySelector('title')?.textContent || 'Inline SVG';
        addImage(dataUrl, alt, null);
      } catch {
        // ignore serialization errors
      }
    }

    // 4. SVG image and use elements
    if (tagName === 'image') {
      const href = el.getAttribute('href') || el.getAttribute('xlink:href');
      addImage(href, '', null);
    }

    // 5. Embedded objects / embeds
    if (tagName === 'object') {
      const data = el.getAttribute('data');
      const type = el.getAttribute('type');
      if (data && (type?.startsWith('image/') || data.match(/\.(png|jpg|jpeg|gif|svg|webp|avif|bmp|ico)(\?.*)?$/i))) {
        addImage(data, '', null);
      }
    }
    if (tagName === 'embed') {
      const src = el.getAttribute('src');
      const type = el.getAttribute('type');
      if (src && (type?.startsWith('image/') || src.match(/\.(png|jpg|jpeg|gif|svg|webp|avif|bmp|ico)(\?.*)?$/i))) {
        addImage(src, '', null);
      }
    }

    // 6. CSS Background images
    try {
      const style = window.getComputedStyle(el);
      const bgImage = style.backgroundImage || style.background;
      if (bgImage && bgImage !== 'none') {
        const urlMatches = bgImage.matchAll(/url\(\s*['"]?([^'")]+)['"]?\s*\)/g);
        for (const match of urlMatches) {
          const url = match[1];
          if (url && !url.startsWith('data:font/')) {
            addImage(url, '', null);
          }
        }
      }
    } catch {
      // ignore style reading errors
    }
  }

  // 7. Stylesheet parsing for background images not active in current viewport elements
  try {
    for (const sheet of document.styleSheets) {
      try {
        const rules = sheet.cssRules || sheet.rules;
        if (rules) {
          for (const rule of rules) {
            if (rule.style) {
              const bg = rule.style.backgroundImage || rule.style.background;
              if (bg && bg !== 'none') {
                const urlMatches = bg.matchAll(/url\(\s*['"]?([^'")]+)['"]?\s*\)/g);
                for (const match of urlMatches) {
                  const url = match[1];
                  if (url && !url.startsWith('data:font/')) {
                    addImage(url, '', null);
                  }
                }
              }
            }
          }
        }
      } catch {
        // ignore cross-origin stylesheet errors
      }
    }
  } catch {
    // ignore styleSheet access errors
  }

  // 8. Link elements (icons)
  const links = document.querySelectorAll('link[rel*="icon"]');
  for (const link of links) {
    const href = link.getAttribute('href');
    addImage(href, 'Favicon', null);
  }

  // 9. Meta elements (social sharing images)
  const metas = document.querySelectorAll('meta[property*="image"], meta[name*="image"]');
  for (const meta of metas) {
    const content = meta.getAttribute('content');
    if (content) {
      addImage(content, 'Meta Image', null);
    }
  }

  return images;
}

/**
 * Resolves the image size in KB and width in pixels.
 * Uses local base64 computation for data URLs,
 * fallback HTTP requests for size, and browser DOM evaluation for dimension.
 * 
 * @param {import('playwright').BrowserContext} context - Playwright Browser Context.
 * @param {import('playwright').Page} page - Playwright Page object.
 * @param {string} imageUrl - Absolute URL of the image.
 * @param {number|null} existingWidth - Existing width if already known.
 * @param {number|null} existingSize - Existing size if already known from headers.
 * @returns {Promise<{sizeKB: number|null, width: number|null}>} Resolved metadata.
 */
async function resolveImageMetadata(context, page, imageUrl, existingWidth, existingSize) {
  let sizeKB = existingSize;
  let width = existingWidth;

  // 1. Resolve size for data URLs locally
  if (imageUrl.startsWith('data:')) {
    try {
      let sizeBytes = 0;
      const match = imageUrl.match(/^data:([^;]+);(base64),(.*)$/);
      if (match) {
        const base64Str = match[3];
        const padding = (base64Str.endsWith('==') ? 2 : (base64Str.endsWith('=') ? 1 : 0));
        sizeBytes = (base64Str.length * 3) / 4 - padding;
      } else {
        const dataPart = imageUrl.split(',')[1] || '';
        let decoded = dataPart;
        try {
          decoded = decodeURIComponent(dataPart);
        } catch {
          // fallback if URI malformed
        }
        sizeBytes = Buffer.byteLength(decoded, 'utf-8');
      }
      sizeKB = Math.round((sizeBytes / 1024) * 100) / 100;
    } catch {
      // ignore parsing errors
    }

    if (!width) {
      const dimensions = await page.evaluate(async (url) => {
        return new Promise((resolve) => {
          const img = new Image();
          const timer = setTimeout(() => {
            img.onload = null;
            img.onerror = null;
            resolve({ width: null });
          }, 2000);
          img.onload = () => {
            clearTimeout(timer);
            resolve({ width: img.naturalWidth || null });
          };
          img.onerror = () => {
            clearTimeout(timer);
            resolve({ width: null });
          };
          img.src = url;
        });
      }, imageUrl).catch(() => ({ width: null }));

      width = dimensions.width;
    }

    return { sizeKB, width };
  }

  // 2. Resolve width in page context if missing
  if (!width) {
    try {
      const dimensions = await page.evaluate(async (url) => {
        return new Promise((resolve) => {
          const img = new Image();
          const timer = setTimeout(() => {
            img.onload = null;
            img.onerror = null;
            resolve({ width: null });
          }, 3000);
          img.onload = () => {
            clearTimeout(timer);
            resolve({ width: img.naturalWidth || null });
          };
          img.onerror = () => {
            clearTimeout(timer);
            resolve({ width: null });
          };
          img.src = url;
        });
      }, imageUrl).catch(() => ({ width: null }));

      width = dimensions.width;
    } catch {
      // ignore evaluate errors
    }
  }

  // 3. Fetch missing size via Playwright request API
  if (sizeKB === null || sizeKB === undefined) {
    try {
      const res = await context.request.get(imageUrl, { timeout: 5000 }).catch(() => null);
      if (res && res.ok()) {
        const headers = res.headers();
        let size = null;
        if (headers['content-length']) {
          size = Number(headers['content-length']);
        }
        if (size === null || Number.isNaN(size)) {
          const body = await res.body();
          size = body ? body.byteLength : null;
        }
        if (size !== null && !Number.isNaN(size)) {
          sizeKB = Math.round((size / 1024) * 100) / 100;
        }
      }
    } catch {
      // ignore fetch errors
    }
  }

  return { sizeKB, width };
}

/**
 * Concurrency helper to limit parallel executions.
 * 
 * @template T
 * @param {(() => Promise<T>)[]} tasks - Array of thunk functions returning promises.
 * @param {number} limit - Max concurrent promises.
 * @returns {Promise<T[]>} Resolved results of all tasks.
 */
async function limitConcurrency(tasks, limit) {
  const results = [];
  const executing = new Set();
  for (const task of tasks) {
    const p = Promise.resolve().then(() => task());
    results.push(p);
    executing.add(p);
    const clean = () => executing.delete(p);
    p.then(clean, clean);
    if (executing.size >= limit) {
      await Promise.race(executing);
    }
  }
  return Promise.all(results);
}

/**
 * Crawls a target URL and its same-origin pages, extracting images and metadata.
 * 
 * @param {string} startUrl - Initial URL.
 * @param {boolean} singlePage - Whether to limit crawling to the single page.
 * @returns {Promise<any[]>} List of all extracted images with their metadata.
 */
async function crawl(startUrl, singlePage) {
  const browser = await chromium.launch({ headless: true });
  const imageSizes = new Map();

  try {
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

    const page = await context.newPage();

    // Listen to response sizes during navigation to optimize size retrieval
    page.on('response', async (response) => {
      try {
        const rawUrl = response.url();
        const normUrl = normalizeUrl(rawUrl, startUrl);
        if (!normUrl) return;

        const request = response.request();
        const resourceType = request.resourceType();
        const headers = response.headers();
        const contentType = headers['content-type'] || '';

        if (resourceType === 'image' || contentType.startsWith('image/')) {
          let size = null;
          if (headers['content-length']) {
            const parsed = Number(headers['content-length']);
            if (!Number.isNaN(parsed)) {
              size = parsed;
            }
          }
          if (size === null) {
            const body = await response.body().catch(() => null);
            size = body ? body.byteLength : null;
          }
          if (size !== null) {
            imageSizes.set(normUrl, Math.round((size / 1024) * 100) / 100);
          }
        }
      } catch {
        // ignore response intercept errors
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
        const { routes } = extractRoutes(html, currentUrl);

        // Scroll to load lazy images before extracting DOM
        await autoScroll(page);

        // Extract all raw image locations from page DOM and styles
        const pageImageRawList = await page.evaluate(extractAllImagesFromPage);

        // Group by normalized image URL to deduplicate redundant records on the same route page
        const uniqueImageMap = new Map();
        for (const rawImg of pageImageRawList) {
          const normImgUrl = normalizeUrl(rawImg.src, currentUrl);
          if (!normImgUrl) continue;

          // Exclude stylesheet and webfont files that may match URL pattern but are not images
          if (normImgUrl.match(/\.(css|woff|woff2|ttf|otf|eot)(\?.*)?$/i)) {
            continue;
          }

          if (uniqueImageMap.has(normImgUrl)) {
            const existing = uniqueImageMap.get(normImgUrl);
            const newAlt = (rawImg.alt && rawImg.alt.length > existing.alt.length) ? rawImg.alt : existing.alt;
            const newWidth = rawImg.width ? Math.max(existing.width || 0, rawImg.width) : existing.width;
            uniqueImageMap.set(normImgUrl, { alt: newAlt, width: newWidth });
          } else {
            uniqueImageMap.set(normImgUrl, { alt: rawImg.alt, width: rawImg.width });
          }
        }

        // Resolve missing metadata (sizes and dimensions) concurrently
        const resolutionTasks = Array.from(uniqueImageMap.entries()).map(([normImgUrl, details]) => {
          return async () => {
            const cachedSize = imageSizes.get(normImgUrl) ?? null;
            const resolved = await resolveImageMetadata(context, page, normImgUrl, details.width, cachedSize);
            return {
              route: currentUrl,
              image: normImgUrl,
              alt: details.alt || '',
              size: resolved.sizeKB,
              width: resolved.width
            };
          };
        });

        // Use limit of 10 concurrent requests to resolve metadata quickly and safely
        const resolvedImages = await limitConcurrency(resolutionTasks, 10);
        allImages.push(...resolvedImages);

        // Queue subpages if in recursive mode
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

    return allImages;
  } finally {
    // Section 6.1: Always close resources to prevent process leaks
    await browser.close();
  }
}

/**
 * Writes the collected images data into an Excel spreadsheet.
 * 
 * @param {any[]} allImages - List of resolved images.
 * @param {string} hostname - Target website hostname for output file name.
 * @returns {string} Path to the created Excel spreadsheet.
 */
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

/**
 * Parses the domain hostname from a URL.
 * 
 * @param {string} urlString - URL string.
 * @returns {string} Parsed hostname or default name.
 */
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
