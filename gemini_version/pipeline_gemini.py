#!/usr/bin/env python3
"""
GobLab Pipeline — Gemini version

Phase 1  SCRAPE   — Google News RSS (20 terms, last 12 months) + direct Chilean
                    news RSS feeds + expanded institutional sites + MercadoPublico
                    procurement tenders → candidate URLs
Phase 1b RESOLVE  — Playwright (async, shared browser) follows Google News JS
                    redirects to get real article URLs
Phase 2  TRIAGE   — Gemini reads each article's full text; keeps only specific
                    Chilean public-sector AI/algorithmic systems; explicitly
                    rejects: education programs about AI, dead links, articles
                    about GobLab's own repository, private-sector-only systems
Phase 3  GROUP    — Clusters URLs that describe the same project
Phase 4  GENERATE — Produces one Word ficha per project cluster

Model mapping vs. Claude version:
  claude-haiku-4-5  → gemini-2.5-flash  (triage + grouping)
  claude-opus-4-8   → gemini-2.5-pro    (extraction)

Note: Anthropic prompt caching is not used in this version. Gemini context
caching requires a separate setup and minimum token counts; adding it is left
as a future optimization.

Usage:
    python3 pipeline_gemini.py
    python3 pipeline_gemini.py --max-candidates 50
    python3 pipeline_gemini.py --output mis_fichas.docx

Requires: GOOGLE_API_KEY environment variable
"""

import sys
import os
import csv
import json
import time
import asyncio
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import feedparser
import requests
from bs4 import BeautifulSoup
from google import genai
from google.genai import types

from generate_ficha_gemini import fetch_url, extract_ficha_fields, build_docx, OUTPUT_DIR, SYSTEM_PROMPT

# ── MODELS ────────────────────────────────────────────────────────────────────

TRIAGE_MODEL  = "gemini-2.5-flash"
EXTRACT_MODEL = "gemini-2.5-pro"

# ── SCRAPING CONFIG ───────────────────────────────────────────────────────────

_ONE_YEAR_AGO = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")

def _news_terms_with_date():
    terms = [
        "inteligencia artificial piloto Chile gobierno",
        "IA implementación gobierno Chile",
        "sistema IA piloto Chile",
        "algoritmos sector público Chile",
        "automatización decisiones gobierno Chile",
        "machine learning sector público Chile",
        "analítica predictiva gobierno Chile",
        "modelos predictivos sector público Chile",
        "licitación inteligencia artificial Chile gobierno",
        "software IA ministerio Chile",
        "plataforma digital gobierno Chile algoritmo",
        "sistema predictivo servicio público Chile",
        "reconocimiento facial Chile gobierno",
        "IA fiscalización Chile servicio público",
        "inteligencia artificial salud Chile ministerio",
        "inteligencia artificial educación Chile ministerio",
        "modelo predictivo Chile ministerio",
        "herramienta IA institución pública Chile",
        "artificial intelligence Chile government public sector",
        "machine learning Chile ministry algorithm",
    ]
    return [f"{t} after:{_ONE_YEAR_AGO}" for t in terms]

GOOGLE_NEWS_TERMS = _news_terms_with_date()

DIRECT_NEWS_FEEDS = [
    {"name": "Radio U. Chile",  "url": "https://radio.uchile.cl/feed/"},
    {"name": "La Tercera",      "url": "https://www.latercera.com/rss/"},
    {"name": "CIPER Chile",     "url": "https://www.ciperchile.cl/feed/"},
    {"name": "The Clinic",      "url": "https://www.theclinic.cl/feed/"},
    {"name": "ANID RSS",        "url": "https://anid.cl/feed/"},
    {"name": "ChileCompra",     "url": "https://www.chilecompra.cl/feed/"},
]

MERCADOPUBLICO_TERMS = [
    "inteligencia artificial",
    "machine learning",
    "algoritmo",
    "modelo predictivo",
]

INSTITUTIONAL_SOURCES = [
    {
        "name": "Gobierno de Chile — Noticias",
        "url":  "https://www.gob.cl/noticias/",
        "base": "https://www.gob.cl",
        "article_only": True,
    },
    {
        "name": "División Gobierno Digital",
        "url":  "https://digital.gob.cl/media/noticias/",
        "base": "https://digital.gob.cl",
        "article_only": True,
    },
    {
        "name": "Ministerio de Ciencia — Noticias",
        "url":  "https://www.minciencia.gob.cl/noticias/",
        "base": "https://www.minciencia.gob.cl",
        "article_only": True,
    },
    {
        "name": "ANID — Noticias",
        "url":  "https://www.anid.cl/noticias/",
        "base": "https://www.anid.cl",
        "article_only": True,
    },
    {
        "name": "ANID — Búsqueda IA",
        "url":  "https://www.anid.cl/?s=inteligencia+artificial",
        "base": "https://www.anid.cl",
        "article_only": True,
    },
    {
        "name": "TransformacionPublica — Blog",
        "url":  "https://transformacionpublica.cl/blog/",
        "base": "https://transformacionpublica.cl",
        "article_only": True,
    },
    {
        "name": "Ministerio de Economía",
        "url":  "https://www.economia.gob.cl/category/noticias",
        "base": "https://www.economia.gob.cl",
        "article_only": True,
    },
    {
        "name": "Servicio Civil",
        "url":  "https://www.serviciocivil.cl/noticias/",
        "base": "https://www.serviciocivil.cl",
        "article_only": True,
    },
    {
        "name": "Laboratorio de Gobierno",
        "url":  "https://www.lab.gob.cl/noticias",
        "base": "https://www.lab.gob.cl",
        "article_only": True,
    },
    {
        "name": "DIPRES — Balance de Gestión Integral",
        "url":  "https://www.dipres.gob.cl/597/w3-propertyvalue-15160.html",
        "base": "https://www.dipres.gob.cl",
        "article_only": False,
    },
]

HEADERS = {
    "User-Agent":      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-CL,es;q=0.9,en;q=0.8",
}

SKIP_EXTENSIONS = {".css", ".js", ".png", ".jpg", ".jpeg", ".gif", ".svg", ".ico", ".woff", ".ttf", ".xml"}
SKIP_DOMAINS    = {"facebook.com", "twitter.com", "instagram.com", "linkedin.com", "youtube.com", "youtu.be"}

# ── URL MEMORY ────────────────────────────────────────────────────────────────

SEEN_URLS_PATH = Path("seen_urls.json")


def load_seen_urls() -> set:
    if SEEN_URLS_PATH.exists():
        with open(SEEN_URLS_PATH) as f:
            return set(json.load(f))
    return set()


def save_seen_urls(seen: set) -> None:
    with open(SEEN_URLS_PATH, "w") as f:
        json.dump(sorted(seen), f, indent=2, ensure_ascii=False)


# ── EXISTING DATABASE ─────────────────────────────────────────────────────────

DEFAULT_DB_PATH = Path(
    "given materials/Organización Casos Repositorio - Proyectos Repositorio.csv"
)


def load_algorithm_database(csv_path: str) -> list[dict]:
    path = Path(csv_path)
    if not path.exists():
        print(f"  WARNING: Database file not found: {csv_path}")
        return []
    try:
        with open(path, encoding="utf-8-sig") as f:
            reader  = csv.reader(f)
            next(reader)
            headers = next(reader)
            rows    = []
            for row in reader:
                if row:
                    padded = row + [""] * max(0, len(headers) - len(row))
                    rows.append(dict(zip(headers, padded)))
        print(f"  Base de datos cargada: {len(rows)} algoritmos existentes")
        return rows
    except Exception as e:
        print(f"  WARNING: Could not read database CSV: {e}")
        return []


def build_db_summary(existing: list[dict]) -> str:
    if not existing:
        return ""
    lines = []
    for row in existing:
        name = row.get("Título", "").strip().lstrip("-").strip()
        inst_raw = row.get("Institución Pública", "").strip()
        inst = inst_raw.split("\n")[0].lstrip("-").strip()
        if name:
            lines.append(f"- {name}" + (f" ({inst})" if inst else ""))
    if not lines:
        return ""
    return (
        "\n\nSISTEMAS YA EN LA BASE DE DATOS — marca is_duplicate=true si el artículo "
        "trata sobre alguno de estos (misma institución + mismo sistema):\n"
        + "\n".join(lines)
    )


# ── MERCADOPUBLICO SCRAPER ────────────────────────────────────────────────────

def scrape_mercadopublico() -> list[dict]:
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("  WARNING: playwright not installed — skipping MercadoPublico")
        return []

    import re as _re

    async def _search_all():
        results = []
        seen_codes: set = set()

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page    = await browser.new_page()

            for term in MERCADOPUBLICO_TERMS:
                try:
                    await page.goto("https://www.mercadopublico.cl/Home", timeout=30000)
                    await asyncio.sleep(1)
                    await page.fill("#txtBuscar", term)
                    await page.keyboard.press("Enter")
                    await page.wait_for_load_state("networkidle", timeout=20000)
                    await asyncio.sleep(3)

                    search_frame = next(
                        (f for f in page.frames
                         if "BuscarLicitacion" in f.url and "?" in f.url),
                        None,
                    )
                    if not search_frame:
                        print(f"    MercadoPublico \"{term}\": no se encontró iframe de resultados")
                        continue

                    html      = await search_frame.evaluate("() => document.body.innerHTML")
                    urls_found = _re.findall(
                        r"verFicha\('(https?://www\.mercadopublico\.cl/Procurement/[^']+)'\)",
                        html,
                    )
                    new_count = 0
                    for url in urls_found:
                        code = url.split("idlicitacion=")[-1] if "idlicitacion=" in url else url
                        if code not in seen_codes:
                            seen_codes.add(code)
                            results.append({
                                "url":    url,
                                "title":  "",
                                "source": f'MercadoPublico: "{term}"',
                            })
                            new_count += 1
                    print(f"    MercadoPublico \"{term}\": {new_count} licitaciones")

                except Exception as e:
                    print(f"    ERROR MercadoPublico \"{term}\": {e}")

            await browser.close()

        return results

    return asyncio.run(_search_all())


# ── PHASE 1: SCRAPE ───────────────────────────────────────────────────────────

def scrape_google_news(term: str) -> list[dict]:
    q   = requests.utils.quote(term)
    url = f"https://news.google.com/rss/search?q={q}&hl=es-419&gl=CL&ceid=CL:es-419"
    try:
        r    = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
        feed = feedparser.parse(r.content)
        results = []
        for entry in feed.entries:
            link = entry.get("link", "")
            if link.startswith("http"):
                results.append({
                    "url":    link,
                    "title":  entry.get("title", ""),
                    "source": f'Google News: "{term}"',
                })
        return results
    except Exception as e:
        print(f"    ERROR RSS '{term}': {e}")
        return []


def scrape_direct_feeds() -> list[dict]:
    results = []
    for feed_info in DIRECT_NEWS_FEEDS:
        try:
            r    = requests.get(feed_info["url"], headers=HEADERS, timeout=15)
            r.raise_for_status()
            feed = feedparser.parse(r.content)
            count = 0
            for entry in feed.entries:
                link = entry.get("link", "")
                if not link.startswith("http") or "news.google.com" in link:
                    continue
                results.append({
                    "url":    link,
                    "title":  entry.get("title", ""),
                    "source": f"RSS: {feed_info['name']}",
                })
                count += 1
            print(f"    {feed_info['name']}: {count} artículos")
        except Exception as e:
            print(f"    ERROR {feed_info['name']}: {e}")
    return results


def resolve_href(href, base):
    if not href:
        return None
    href = href.strip()
    if href.startswith(("mailto:", "javascript:", "#", "tel:")):
        return None
    if href.startswith("http"):
        return href
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("/"):
        return base.rstrip("/") + href
    return None


def is_article_link(url):
    from urllib.parse import urlparse
    path  = urlparse(url).path.rstrip("/")
    if not path:
        return False
    slug  = path.split("/")[-1]
    words = [w for w in slug.replace("-", " ").replace("_", " ").split() if len(w) > 1]
    return len(words) >= 5


def scrape_institutional(source) -> list[dict]:
    try:
        r = requests.get(source["url"], headers=HEADERS, timeout=20, allow_redirects=True)
        r.raise_for_status()
        soup         = BeautifulSoup(r.text, "html.parser")
        results      = []
        seen         = set()
        article_only = source.get("article_only", False)

        for tag in soup.find_all("a", href=True):
            href     = tag.get("href", "")
            text     = tag.get_text(strip=True)
            resolved = resolve_href(href, source["base"])
            if not resolved:
                continue
            if any(resolved.lower().endswith(ext) for ext in SKIP_EXTENSIONS):
                continue
            if any(d in resolved for d in SKIP_DOMAINS):
                continue
            if article_only and not is_article_link(resolved):
                continue
            if resolved in seen:
                continue
            seen.add(resolved)
            results.append({"url": resolved, "title": text, "source": source["name"]})
        return results
    except Exception as e:
        print(f"    ERROR {source['name']}: {e}")
        return []


# ── PHASE 1b: RESOLVE GOOGLE NEWS REDIRECTS ───────────────────────────────────

def resolve_gnews_urls(gnews_candidates: list[dict], concurrency: int = 8) -> list[dict]:
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("  WARNING: playwright not installed — skipping Google News URL resolution")
        print("  Run: pip3 install playwright && playwright install chromium")
        return []

    print(f"\n  Resolviendo {len(gnews_candidates)} URLs de Google News...")

    async def _resolve_all(candidates):
        sem = asyncio.Semaphore(concurrency)

        async def resolve_one(candidate):
            async with sem:
                url = candidate["url"]
                try:
                    page = await browser.new_page()
                    await page.goto(url, wait_until="commit", timeout=15000)
                    await asyncio.sleep(2)
                    final = page.url
                    await page.close()
                    if "news.google.com" not in final:
                        return {**candidate, "url": final}
                    return None
                except Exception:
                    try:
                        await page.close()
                    except Exception:
                        pass
                    return None

        async with async_playwright() as p:
            browser  = await p.chromium.launch(headless=True)
            results  = await asyncio.gather(*[resolve_one(c) for c in candidates])
            await browser.close()

        return [r for r in results if r]

    resolved = asyncio.run(_resolve_all(gnews_candidates))

    success = len(resolved)
    total   = len(gnews_candidates)
    print(f"  → {success}/{total} URLs de Google News resueltas\n")
    return resolved


def phase_scrape(max_candidates: int | None, seen_urls: set | None = None) -> list[dict]:
    print("\n" + "═" * 65)
    print("  FASE 1 — SCRAPING")
    print(f"  Buscando artículos desde {_ONE_YEAR_AGO} en adelante")
    print("═" * 65)

    all_candidates = []

    print(f"\n  [Google News RSS — {len(GOOGLE_NEWS_TERMS)} términos de búsqueda]")
    gnews_raw = []
    for term in GOOGLE_NEWS_TERMS:
        results = scrape_google_news(term)
        display = term.split(" after:")[0]
        print(f"    \"{display}\": {len(results)} artículos")
        gnews_raw.extend(results)

    seen_gnews  = set()
    gnews_unique = []
    for c in gnews_raw:
        if c["url"] not in seen_gnews:
            seen_gnews.add(c["url"])
            gnews_unique.append(c)
    print(f"  → {len(gnews_unique)} URLs únicas de Google News")

    resolved = resolve_gnews_urls(gnews_unique)
    all_candidates.extend(resolved)

    print("  [Feeds RSS directos]")
    all_candidates.extend(scrape_direct_feeds())

    print("\n  [Sitios institucionales]")
    for source in INSTITUTIONAL_SOURCES:
        results = scrape_institutional(source)
        print(f"    {source['name']}: {len(results)} links")
        all_candidates.extend(results)

    print("\n  [MercadoPublico — licitaciones IA]")
    all_candidates.extend(scrape_mercadopublico())

    seen   = set()
    unique = []
    for c in all_candidates:
        if c["url"] not in seen:
            seen.add(c["url"])
            unique.append(c)

    print(f"\n  → {len(unique)} URLs únicas en total")

    if seen_urls:
        before  = len(unique)
        unique  = [c for c in unique if c["url"] not in seen_urls]
        skipped = before - len(unique)
        if skipped:
            print(f"  → {skipped} omitidas (ya procesadas anteriormente)")

    print(f"  → {len(unique)} URLs nuevas para triage")

    if max_candidates and len(unique) > max_candidates:
        unique = unique[:max_candidates]
        print(f"  → Limitado a {max_candidates} (--max-candidates)")

    return unique


# ── PHASE 2: TRIAGE ───────────────────────────────────────────────────────────

def parse_json(text: str) -> dict:
    """Extract and parse JSON from model output that may contain preamble text."""
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()
    if not text.startswith("{"):
        start = text.find("{")
        end   = text.rfind("}") + 1
        if start != -1 and end > 0:
            text = text[start:end]
    return json.loads(text)


TRIAGE_SYSTEM_BASE = """\
Eres un asistente de investigación para el Repositorio de Algoritmos Públicos de GobLab UAI \
(Universidad Adolfo Ibáñez, Chile).
Evalúa si un artículo describe un sistema de IA o algoritmo ESPECÍFICO que está siendo usado \
u oficialmente planificado por una institución pública chilena.

Responde ÚNICAMENTE con JSON válido con estas claves exactas:
{
  "relevant": true o false,
  "reason": "una oración explicando la decisión",
  "system_name": "nombre del sistema/proyecto, o null",
  "institution": "nombre de la institución pública chilena, o null",
  "is_duplicate": true o false,
  "duplicate_of": "nombre del sistema existente en la base de datos, o null"
}

relevant = true SOLO si se cumplen TODOS estos criterios:
- Describe un sistema, herramienta, o proyecto ESPECÍFICO con nombre o función identificable
  (no tendencias generales de IA, no estudios, no rankings, no conferencias)
- Involucra al sector público chileno (gobierno, ministerios, servicios públicos, municipios)
- El sistema está activo, en piloto oficial, o en planificación formal anunciada
- El artículo tiene suficiente contenido textual para evaluar el sistema

relevant = false en CUALQUIERA de estos casos:
- El artículo habla de un programa de formación, curso, taller, o educación sobre IA
  (ej. "X institución capacita a funcionarios en IA" — eso es educación, no un sistema)
- El artículo es sobre el Repositorio de Algoritmos Públicos de GobLab UAI en sí mismo
- La URL retorna un error 404 o el contenido dice "página no encontrada"
- Involucra solo sector privado sin contrato/uso/mandato público
- Habla de otro país y no del sector público de Chile específicamente
- La mención de Chile o de IA es incidental o de relleno
- Es principalmente video, ranking, opinión, o entretenimiento sin descripción técnica

is_duplicate = true si el sistema ya aparece en la lista de sistemas existentes.\
"""

GROUP_SYSTEM = """\
Eres un asistente que agrupa artículos sobre algoritmos del sector público chileno.
Se te entregará una lista numerada de artículos, cada uno con el nombre del sistema \
detectado y la institución.
Tu tarea: identificar cuáles artículos hablan del MISMO proyecto o sistema, \
aunque usen nombres ligeramente distintos.

Responde ÚNICAMENTE con JSON válido:
{
  "groups": [
    {
      "canonical_name":        "nombre canónico del sistema",
      "canonical_institution": "institución canónica",
      "indices":               [0, 3, 7]
    }
  ]
}

Reglas:
- Cada índice debe aparecer exactamente una vez.
- Si un artículo no tiene par claro, ponlo solo en su propio grupo.
- Usa el nombre e institución más oficial/completo como canónico.\
"""


def triage_one(candidate: dict, client: genai.Client, db_suffix: str) -> dict:
    url  = candidate["url"]
    text = fetch_url(url)

    if (len(text) < 300 or text.startswith("[ERROR") or
            any(p in text.lower() for p in ["página no encontrada", "page not found",
                                            "404 not found", "error 404"])):
        return {
            "url": url, "relevant": False,
            "reason": "Contenido insuficiente, error al obtener la página, o 404",
            "system_name": None, "institution": None,
            "is_duplicate": False, "duplicate_of": None,
            "orig_source": candidate.get("source", ""),
        }

    content = f"URL: {url}\n\nContenido:\n{text[:4000]}"

    try:
        response = client.models.generate_content(
            model=TRIAGE_MODEL,
            config=types.GenerateContentConfig(
                system_instruction=TRIAGE_SYSTEM_BASE + db_suffix,
                max_output_tokens=600,
                response_mime_type="application/json",
            ),
            contents=content,
        )
        result = parse_json(response.text)
    except Exception as e:
        result = {
            "relevant": False, "reason": f"Error en triage: {e}",
            "system_name": None, "institution": None,
            "is_duplicate": False, "duplicate_of": None,
        }

    result["url"]         = url
    result["orig_source"] = candidate.get("source", "")
    return result


def phase_triage(candidates: list[dict], client: genai.Client, db_suffix: str) -> list[dict]:
    print("\n" + "═" * 65)
    print("  FASE 2 — TRIAGE")
    print(f"  Gemini ({TRIAGE_MODEL}) lee el texto completo de cada artículo")
    print("═" * 65)
    print(f"\n  Analizando {len(candidates)} URLs...\n")

    results = [None] * len(candidates)
    done    = 0

    with ThreadPoolExecutor(max_workers=6) as pool:
        future_to_idx = {
            pool.submit(triage_one, c, client, db_suffix): i
            for i, c in enumerate(candidates)
        }
        for future in as_completed(future_to_idx):
            idx          = future_to_idx[future]
            result       = future.result()
            results[idx] = result
            done        += 1

            if result.get("is_duplicate"):
                mark  = "⟳"
                label = f"DUPLICADO de: {result.get('duplicate_of') or '?'}"
            elif result.get("relevant"):
                mark  = "✓"
                label = f"{result.get('system_name') or '—'} | {result.get('institution') or '—'}"
            else:
                mark  = "✗"
                label = result.get("reason", "")[:60]

            print(f"  [{done:>3}/{len(candidates)}] {mark}  {result['url'][:60]}")
            if result.get("relevant") or result.get("is_duplicate"):
                print(f"             → {label}")

    relevant = [r for r in results if r and r.get("relevant") and not r.get("is_duplicate")]
    dupes    = [r for r in results if r and r.get("is_duplicate")]

    print(f"\n  → {len(relevant)} relevantes | {len(dupes)} duplicados descartados")
    return relevant


# ── PHASE 3: GROUP ────────────────────────────────────────────────────────────

def phase_group(relevant: list[dict], client: genai.Client) -> list[list[dict]]:
    print("\n" + "═" * 65)
    print("  FASE 3 — AGRUPACIÓN POR PROYECTO")
    print("═" * 65)

    if not relevant:
        return []

    if len(relevant) == 1:
        r = relevant[0]
        r["canonical_name"]        = r.get("system_name")
        r["canonical_institution"] = r.get("institution")
        return [[r]]

    lines = "\n".join(
        f"[{i}] Sistema: {r.get('system_name') or 'desconocido'} | "
        f"Institución: {r.get('institution') or 'desconocida'} | "
        f"URL: {r['url']}"
        for i, r in enumerate(relevant)
    )
    print(f"\n  Agrupando {len(relevant)} artículos relevantes...\n")

    try:
        response = client.models.generate_content(
            model=TRIAGE_MODEL,
            config=types.GenerateContentConfig(
                system_instruction=GROUP_SYSTEM,
                max_output_tokens=2000,
                response_mime_type="application/json",
            ),
            contents=f"Agrupa estos artículos:\n\n{lines}",
        )
        parsed = parse_json(response.text)

        groups = []
        for g in parsed.get("groups", []):
            group = []
            for idx in g.get("indices", []):
                if 0 <= idx < len(relevant):
                    item = dict(relevant[idx])
                    item["canonical_name"]        = g.get("canonical_name")
                    item["canonical_institution"] = g.get("canonical_institution")
                    group.append(item)
            if group:
                groups.append(group)

    except Exception as e:
        print(f"  ERROR al agrupar: {e} — cada URL como proyecto independiente")
        groups = []
        for r in relevant:
            item = dict(r)
            item["canonical_name"]        = r.get("system_name")
            item["canonical_institution"] = r.get("institution")
            groups.append([item])

    print(f"  → {len(groups)} proyecto(s) único(s):\n")
    for i, group in enumerate(groups, 1):
        name = group[0].get("canonical_name") or "?"
        inst = group[0].get("canonical_institution") or "?"
        print(f"  [{i}] {name}  |  {inst}  ({len(group)} fuente(s))")
        for item in group:
            print(f"       {item['url']}")

    return groups


# ── PHASE 4: GENERATE FICHAS ──────────────────────────────────────────────────

def phase_generate(groups: list[list[dict]], client: genai.Client, output_path: Path):
    print("\n" + "═" * 65)
    print("  FASE 4 — GENERACIÓN DE FICHAS")
    print(f"  Modelo: {EXTRACT_MODEL}")
    print("═" * 65)

    fichas     = []
    source_map = []

    for i, group in enumerate(groups, 1):
        name = group[0].get("canonical_name") or f"Proyecto {i}"
        print(f"\n  [{i}/{len(groups)}] {name}")

        blocks = []
        for item in group:
            print(f"    Fetching {item['url'][:70]}...")
            text = fetch_url(item["url"])
            if len(text) >= 300 and not text.startswith("[ERROR"):
                blocks.append({"url": item["url"], "text": text})
            else:
                print(f"    Poco contenido ({len(text)} chars) — omitiendo esta fuente")

        if not blocks:
            print(f"    Sin contenido disponible — omitiendo ficha")
            continue

        print(f"    Enviando {len(blocks)} fuente(s) a Gemini para extracción...")
        ficha = extract_ficha_fields(blocks, client)
        fichas.append(ficha)
        source_map.append(blocks)

        if "_parse_error" not in ficha:
            n = len(ficha.get("fuentes_apa") or [])
            print(f"    → {ficha.get('nombre', '?')} | {ficha.get('institucion_publica', '?')} | {n} cita(s)")

    OUTPUT_DIR.mkdir(exist_ok=True)
    build_docx(fichas, source_map, output_path)


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="GobLab pipeline (Gemini): scrape → triage → group → fichas"
    )
    parser.add_argument("--max-candidates", type=int, default=None,
                        help="Cap URLs sent to triage (for quick tests)")
    parser.add_argument("--output", "-o", default=None,
                        help="Output .docx filename (saved in fichas_output/)")
    parser.add_argument("--db", default=None,
                        help="CSV with existing algorithms for deduplication")
    args = parser.parse_args()

    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        print("ERROR: Set GOOGLE_API_KEY environment variable.")
        sys.exit(1)

    client    = genai.Client(api_key=api_key)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name  = args.output or f"pipeline_{timestamp}.docx"
    out_path  = OUTPUT_DIR / out_name

    print("\n" + "═" * 65)
    print("  GobLab — Pipeline automatizado (Gemini)")
    print(f"  Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"  Salida: fichas_output/{out_name}")
    print("═" * 65)

    db_suffix = ""
    db_path   = args.db or (str(DEFAULT_DB_PATH) if DEFAULT_DB_PATH.exists() else None)
    if db_path:
        print(f"\n  [Base de datos: {db_path}]")
        existing  = load_algorithm_database(db_path)
        db_suffix = build_db_summary(existing)
    else:
        print("\n  (Sin base de datos — colocar CSV en 'given materials/' para deduplicación)")

    t0 = time.time()

    seen_urls  = load_seen_urls()
    if seen_urls:
        print(f"\n  [Memoria de URLs: {len(seen_urls)} URLs previamente procesadas]")

    candidates = phase_scrape(args.max_candidates, seen_urls)
    relevant   = phase_triage(candidates, client, db_suffix)

    new_seen = seen_urls | {c["url"] for c in candidates}
    save_seen_urls(new_seen)
    print(f"\n  [Memoria de URLs actualizada: {len(new_seen)} URLs en total]")

    groups = phase_group(relevant, client)

    if not groups:
        print("\n  No se encontraron proyectos relevantes. Terminando sin generar fichas.")
        sys.exit(0)

    phase_generate(groups, client, out_path)

    elapsed = int(time.time() - t0)
    mins, secs = divmod(elapsed, 60)
    print(f"\n  Tiempo total: {mins}m {secs}s")
    print(f"  Fichas guardadas en: {out_path}\n")


if __name__ == "__main__":
    main()
