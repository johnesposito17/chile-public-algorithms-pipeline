#!/usr/bin/env node
'use strict';

/**
 * GobLab URL Scraper — Step 1
 *
 * Finds candidate URLs about AI / algorithmic systems in Chilean public institutions.
 * Does NOT analyze content — just collects and saves URLs.
 *
 * Sources:
 *   1. Google News RSS — 8 Spanish search terms
 *   2. Institutional sites: anid.cl, mercadopublico.cl, dipres.gob.cl, transformacionpublica.cl
 *
 * Output: candidate_urls_YYYY-MM-DD.txt
 */

const axios  = require('axios');
const cheerio = require('cheerio');
const xml2js  = require('xml2js');
const fs      = require('fs');

// ── CONFIG ────────────────────────────────────────────────────────────────────

const TODAY = new Date().toISOString().split('T')[0]; // YYYY-MM-DD
const OUTPUT_FILE = `candidate_urls_${TODAY}.txt`;

const HTTP_HEADERS = {
  'User-Agent':      'Mozilla/5.0 (compatible; GobLabResearchBot/1.0)',
  'Accept':          'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
  'Accept-Language': 'es-CL,es;q=0.9,en;q=0.8',
};
const TIMEOUT_MS     = 20_000;
const POLITE_DELAY_MS = 1_000; // wait between requests to the same host

// ── GOOGLE NEWS SEARCH TERMS ──────────────────────────────────────────────────

const GOOGLE_NEWS_TERMS = [
  'inteligencia artificial piloto Chile',
  'IA implementación gobierno Chile',
  'sistema IA piloto Chile',
  'algoritmos sector público Chile',
  'automatización decisiones gobierno Chile',
  'machine learning sector público Chile',
  'analítica predictiva gobierno Chile',
  'modelos predictivos sector público Chile',
];

// ── INSTITUTIONAL SOURCES ─────────────────────────────────────────────────────
//
// Each entry has:
//   name     — label printed in output
//   url      — page to fetch
//   base     — used to resolve relative hrefs
//   note     — what we're looking for on this page

const INSTITUTIONAL_SOURCES = [
  {
    name: 'ANID — Noticias',
    url:  'https://www.anid.cl/noticias/',
    base: 'https://www.anid.cl',
    note: 'Noticias generales de ANID',
  },
  {
    name: 'ANID — Noticias IA / ciencia aplicada',
    url:  'https://www.anid.cl/?s=inteligencia+artificial',
    base: 'https://www.anid.cl',
    note: 'Búsqueda en ANID: inteligencia artificial',
  },
  {
    name: 'MercadoPublico — inteligencia artificial',
    url:  'https://www.mercadopublico.cl/Procurement/Modules/RFB/ListSearch.aspx?ModuleType=1&FilterType=1&FilterValue=inteligencia+artificial',
    base: 'https://www.mercadopublico.cl',
    note: 'Licitaciones que mencionan inteligencia artificial',
  },
  {
    name: 'MercadoPublico — machine learning',
    url:  'https://www.mercadopublico.cl/Procurement/Modules/RFB/ListSearch.aspx?ModuleType=1&FilterType=1&FilterValue=machine+learning',
    base: 'https://www.mercadopublico.cl',
    note: 'Licitaciones que mencionan machine learning',
  },
  {
    name: 'MercadoPublico — algoritmo',
    url:  'https://www.mercadopublico.cl/Procurement/Modules/RFB/ListSearch.aspx?ModuleType=1&FilterType=1&FilterValue=algoritmo',
    base: 'https://www.mercadopublico.cl',
    note: 'Licitaciones que mencionan algoritmo',
  },
  {
    name: 'DIPRES — Balance de Gestión Integral',
    url:  'https://www.dipres.gob.cl/597/w3-propertyvalue-15160.html',
    base: 'https://www.dipres.gob.cl',
    note: 'Documentos BGI de servicios públicos',
  },
  {
    name: 'DIPRES — Documentos institucionales',
    url:  'https://www.dipres.gob.cl/597/w3-channel.html',
    base: 'https://www.dipres.gob.cl',
    note: 'Canal de documentos DIPRES',
  },
  {
    name: 'TransformacionPublica.cl',
    url:  'https://www.transformacionpublica.cl/',
    base: 'https://www.transformacionpublica.cl',
    note: 'Noticias y publicaciones sobre transformación pública',
  },
];

// ── HELPERS ───────────────────────────────────────────────────────────────────

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/** Resolve a potentially-relative href to an absolute URL. Returns null if not resolvable. */
function resolveHref(href, base) {
  if (!href) return null;
  href = href.trim();
  if (href.startsWith('mailto:') || href.startsWith('javascript:') || href === '#') return null;
  if (href.startsWith('http://') || href.startsWith('https://')) return href;
  if (href.startsWith('//')) return 'https:' + href;
  if (href.startsWith('/')) return base.replace(/\/$/, '') + href;
  return null; // skip ambiguous relative paths
}

/** Strip tracking parameters and fragments to normalise a URL for dedup. */
function normaliseUrl(url) {
  try {
    const u = new URL(url);
    // Remove common tracking params
    ['utm_source','utm_medium','utm_campaign','utm_term','utm_content',
     'fbclid','gclid','ref','source'].forEach(p => u.searchParams.delete(p));
    u.hash = '';
    return u.toString();
  } catch {
    return url;
  }
}

// ── GOOGLE NEWS RSS ───────────────────────────────────────────────────────────

async function fetchGoogleNewsRss(term) {
  const q      = encodeURIComponent(term);
  const rssUrl = `https://news.google.com/rss/search?q=${q}&hl=es-419&gl=CL&ceid=CL:es-419`;

  console.log(`\n  Término: "${term}"`);
  console.log(`  RSS URL: ${rssUrl}`);

  try {
    const res = await axios.get(rssUrl, { headers: HTTP_HEADERS, timeout: TIMEOUT_MS });

    const parsed  = await xml2js.parseStringPromise(res.data, { explicitArray: true });
    const channel = parsed?.rss?.channel?.[0] ?? {};
    const items   = channel.item ?? [];

    const results = items
      .map(item => {
        const url   = (item.link  ?? [])[0] ?? '';
        const title = (item.title ?? [])[0] ?? '';
        return { url, title, source: `Google News RSS: "${term}"` };
      })
      .filter(r => r.url.startsWith('http'));

    console.log(`  → ${results.length} artículo(s) encontrado(s)`);
    return results;

  } catch (err) {
    console.log(`  → ERROR al obtener RSS: ${err.message}`);
    return [];
  }
}

// ── INSTITUTIONAL SITE SCRAPER ────────────────────────────────────────────────

async function scrapeInstitutionalSite(source) {
  const { name, url, base, note } = source;

  console.log(`\n  Fuente: ${name}`);
  console.log(`  Nota:   ${note}`);
  console.log(`  URL:    ${url}`);

  try {
    const res = await axios.get(url, {
      headers:      HTTP_HEADERS,
      timeout:      TIMEOUT_MS,
      maxRedirects: 5,
    });

    const $ = cheerio.load(res.data);
    const results = [];

    $('a[href]').each((_, el) => {
      const href     = $(el).attr('href');
      const linkText = $(el).text().replace(/\s+/g, ' ').trim();
      const resolved = resolveHref(href, base);

      if (!resolved)                         return; // skip unresolvable
      if (resolved === url)                  return; // skip self-link
      if (resolved.endsWith('.css'))         return; // skip assets
      if (resolved.endsWith('.js'))          return;
      if (resolved.endsWith('.png'))         return;
      if (resolved.endsWith('.jpg'))         return;
      if (resolved.endsWith('.gif'))         return;
      if (resolved.endsWith('.svg'))         return;
      if (resolved.endsWith('.ico'))         return;
      if (resolved.includes('facebook.com')) return; // skip social media
      if (resolved.includes('twitter.com'))  return;
      if (resolved.includes('instagram.com'))return;
      if (resolved.includes('linkedin.com')) return;
      if (resolved.includes('youtube.com'))  return;

      results.push({
        url:    normaliseUrl(resolved),
        title:  linkText || resolved,
        source: name,
      });
    });

    // Deduplicate within this source
    const seen    = new Set();
    const unique  = results.filter(r => {
      if (seen.has(r.url)) return false;
      seen.add(r.url);
      return true;
    });

    console.log(`  → ${unique.length} link(s) encontrado(s) en la página`);
    return unique;

  } catch (err) {
    const status = err.response?.status ?? '';
    console.log(`  → ERROR${status ? ' HTTP ' + status : ''}: ${err.message}`);
    return [];
  }
}

// ── MAIN ──────────────────────────────────────────────────────────────────────

async function main() {
  console.log('='.repeat(65));
  console.log('  GobLab — Buscador de URLs candidatas sobre IA en Chile');
  console.log(`  Fecha de ejecución: ${TODAY}`);
  console.log(`  Archivo de salida:  ${OUTPUT_FILE}`);
  console.log('='.repeat(65));

  const allResults = [];

  // ── STEP 1: Google News RSS ────────────────────────────────────────────────
  console.log('\n\n═══ PASO 1: Google News RSS ═══════════════════════════════════');

  for (const term of GOOGLE_NEWS_TERMS) {
    const results = await fetchGoogleNewsRss(term);
    allResults.push(...results);
    await sleep(POLITE_DELAY_MS);
  }

  const afterRss = allResults.length;
  console.log(`\n  Subtotal Google News: ${afterRss} URLs`);

  // ── STEP 2: Institutional sites ────────────────────────────────────────────
  console.log('\n\n═══ PASO 2: Sitios institucionales ════════════════════════════');

  for (const source of INSTITUTIONAL_SOURCES) {
    const results = await scrapeInstitutionalSite(source);
    allResults.push(...results);
    await sleep(POLITE_DELAY_MS);
  }

  const afterInstitutional = allResults.length - afterRss;
  console.log(`\n  Subtotal sitios institucionales: ${afterInstitutional} URLs`);

  // ── STEP 3: Global deduplication ──────────────────────────────────────────
  const seen   = new Set();
  const unique = allResults.filter(r => {
    const key = normaliseUrl(r.url);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  console.log('\n\n' + '='.repeat(65));
  console.log(`  TOTAL BRUTO:  ${allResults.length} URLs`);
  console.log(`  TOTAL ÚNICO:  ${unique.length} URLs (después de deduplicar)`);
  console.log('='.repeat(65));

  // ── STEP 4: Print all unique URLs ─────────────────────────────────────────
  console.log('\n\n═══ TODAS LAS URLs ÚNICAS ══════════════════════════════════════\n');

  unique.forEach((r, i) => {
    const num   = String(i + 1).padStart(3, ' ');
    const title = r.title && r.title !== r.url
      ? r.title.slice(0, 80) + (r.title.length > 80 ? '…' : '')
      : '';
    console.log(`[${num}] [${r.source}]`);
    console.log(`      ${r.url}`);
    if (title) console.log(`      "${title}"`);
  });

  // ── STEP 5: Save to file ───────────────────────────────────────────────────
  const header = [
    `# GobLab — URLs candidatas sobre IA / sistemas algorítmicos en Chile`,
    `# Generado: ${TODAY}`,
    `# Total URLs únicas: ${unique.length}`,
    `#`,
    `# Formato de cada entrada:`,
    `#   SOURCE: <nombre de la fuente>`,
    `#   URL:    <url>`,
    `#   TITLE:  <texto del enlace o título del artículo>`,
    `#`,
    `# Para analizar una URL con el generador de fichas:`,
    `#   python3 generate_ficha.py <URL>`,
    `# Para analizar desde un archivo de texto:`,
    `#   python3 generate_ficha.py --file ${OUTPUT_FILE}`,
    '',
  ];

  // Also write a plain URL-per-line section for easy piping into generate_ficha.py
  const plainUrlsSection = [
    '',
    '# ── URLS PLANAS (una por línea, para usar con --file) ──────────────',
    ...unique.map(r => r.url),
    '',
  ];

  const detailedSection = [
    '# ── DETALLE COMPLETO ────────────────────────────────────────────────',
    '',
    ...unique.flatMap((r, i) => [
      `# [${i + 1}]`,
      `SOURCE: ${r.source}`,
      `URL:    ${r.url}`,
      r.title && r.title !== r.url ? `TITLE:  ${r.title}` : null,
      '',
    ]).filter(line => line !== null),
  ];

  const fileContent = [...header, ...plainUrlsSection, ...detailedSection].join('\n');
  fs.writeFileSync(OUTPUT_FILE, fileContent, 'utf8');

  console.log(`\n\n✓ ${unique.length} URLs guardadas en: ${OUTPUT_FILE}`);
  console.log(`\n  Para analizar con el generador de fichas:`);
  console.log(`    python3 generate_ficha.py --file ${OUTPUT_FILE}`);
}

main().catch(err => {
  console.error('\n✗ ERROR FATAL:', err.message);
  process.exit(1);
});
