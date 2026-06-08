#!/usr/bin/env python3
"""
GobLab Pipeline — End-to-end automated system

Phase 1  SCRAPE   — Google News RSS + institutional sites → candidate URLs
Phase 2  TRIAGE   — Claude reads each article; keeps only specific Chilean
                    public-sector AI/algorithmic systems; discards general
                    trends, foreign news, sports, video-only pages
Phase 3  GROUP    — Clusters URLs that describe the same project so that
                    multiple sources become multiple APA citations in one ficha
Phase 4  GENERATE — Produces one Word ficha per project cluster

Usage:
    python3 pipeline.py
    python3 pipeline.py --max-candidates 50   # smaller run for testing
    python3 pipeline.py --output mis_fichas.docx
"""

import sys
import os
import json
import time
import argparse
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import feedparser
import requests
from bs4 import BeautifulSoup
import anthropic

from generate_ficha import fetch_url, extract_ficha_fields, build_docx, OUTPUT_DIR

# ── MODELS ────────────────────────────────────────────────────────────────────

TRIAGE_MODEL  = "claude-haiku-4-5-20251001"  # fast + cheap for yes/no filtering
EXTRACT_MODEL = "claude-opus-4-8"            # thorough for final ficha extraction

# ── SCRAPING CONFIG ───────────────────────────────────────────────────────────

GOOGLE_NEWS_TERMS = [
    "inteligencia artificial piloto Chile",
    "IA implementación gobierno Chile",
    "sistema IA piloto Chile",
    "algoritmos sector público Chile",
    "automatización decisiones gobierno Chile",
    "machine learning sector público Chile",
    "analítica predictiva gobierno Chile",
    "modelos predictivos sector público Chile",
    # Additional terms targeting specific government/tender contexts
    "licitación inteligencia artificial Chile gobierno",
    "software IA ministerio Chile",
    "plataforma digital gobierno Chile algoritmo",
    "sistema predictivo servicio público Chile",
]

# MercadoPublico's search pages are JavaScript-rendered — the HTML they return
# is just the navigation shell, not the actual tender listings. Replaced with
# government news portals that serve full HTML article pages.
INSTITUTIONAL_SOURCES = [
    {
        "name": "ANID — Noticias",
        "url":  "https://www.anid.cl/noticias/",
        "base": "https://www.anid.cl",
        "article_only": True,   # filter out nav links; keep only article slugs
    },
    {
        "name": "ANID — Búsqueda IA",
        "url":  "https://www.anid.cl/?s=inteligencia+artificial",
        "base": "https://www.anid.cl",
        "article_only": True,
    },
    {
        "name": "Ministerio de Ciencia — Noticias",
        "url":  "https://www.minciencia.gob.cl/noticias/",
        "base": "https://www.minciencia.gob.cl",
        "article_only": True,
    },
    {
        "name": "Gobierno de Chile — Noticias",
        "url":  "https://www.gob.cl/noticias/",
        "base": "https://www.gob.cl",
        "article_only": True,
    },
    {
        "name": "División Gobierno Digital",
        "url":  "https://digital.gob.cl/noticias/",
        "base": "https://digital.gob.cl",
        "article_only": True,
    },
    {
        "name": "DIPRES — Balance de Gestión Integral",
        "url":  "https://www.dipres.gob.cl/597/w3-propertyvalue-15160.html",
        "base": "https://www.dipres.gob.cl",
        "article_only": False,
    },
    {
        "name": "TransformacionPublica — Blog",
        "url":  "https://transformacionpublica.cl/blog/",
        "base": "https://transformacionpublica.cl",
        "article_only": True,
    },
]

# A realistic browser UA passes CloudFront and most CDN bot-checks
HEADERS = {
    "User-Agent":      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-CL,es;q=0.9,en;q=0.8",
}

SKIP_EXTENSIONS = {".css", ".js", ".png", ".jpg", ".jpeg", ".gif", ".svg", ".ico", ".woff", ".ttf", ".xml"}
SKIP_DOMAINS    = {"facebook.com", "twitter.com", "instagram.com", "linkedin.com", "youtube.com", "youtu.be"}

# ── PHASE 1: SCRAPE ───────────────────────────────────────────────────────────

def scrape_google_news(term):
    q   = requests.utils.quote(term)
    url = f"https://news.google.com/rss/search?q={q}&hl=es-419&gl=CL&ceid=CL:es-419"
    try:
        # feedparser.parse(url) uses its own HTTP client which Google blocks.
        # Fetch the RSS bytes ourselves with our headers, then hand the content
        # to feedparser for parsing.
        r    = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
        feed = feedparser.parse(r.content)
        results = []
        for entry in feed.entries:
            link  = entry.get("link", "")
            title = entry.get("title", "")
            if link.startswith("http"):
                results.append({"url": link, "title": title, "source": f'Google News: "{term}"'})
        return results
    except Exception as e:
        print(f"    ERROR RSS '{term}': {e}")
        return []


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
    """
    Heuristic: news article URLs have long, word-rich slugs like
    /laben-chile-crea-sistema-de-vigilancia-para-el-sector/
    Navigation URLs have short slugs like /convenios/ or /concursos/
    """
    from urllib.parse import urlparse
    path  = urlparse(url).path.rstrip("/")
    if not path:
        return False
    slug  = path.split("/")[-1]
    words = [w for w in slug.replace("-", " ").replace("_", " ").split() if len(w) > 1]
    return len(words) >= 5


def scrape_institutional(source):
    try:
        r = requests.get(source["url"], headers=HEADERS, timeout=20, allow_redirects=True)
        r.raise_for_status()
        soup    = BeautifulSoup(r.text, "html.parser")
        results = []
        seen    = set()
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
            # When article_only is set, skip links that look like nav/category pages
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


def phase_scrape(max_candidates):
    print("\n" + "═" * 65)
    print("  FASE 1 — SCRAPING")
    print("═" * 65)

    candidates = []

    print("\n  [Google News RSS]")
    for term in GOOGLE_NEWS_TERMS:
        results = scrape_google_news(term)
        print(f"    \"{term}\": {len(results)} artículos")
        candidates.extend(results)

    print("\n  [Sitios institucionales]")
    for source in INSTITUTIONAL_SOURCES:
        results = scrape_institutional(source)
        print(f"    {source['name']}: {len(results)} links")
        candidates.extend(results)

    # URL-level dedup
    seen   = set()
    unique = []
    for c in candidates:
        if c["url"] not in seen:
            seen.add(c["url"])
            unique.append(c)

    print(f"\n  → {len(unique)} URLs únicas encontradas")

    if max_candidates and len(unique) > max_candidates:
        unique = unique[:max_candidates]
        print(f"  → Limitado a {max_candidates} para esta ejecución (--max-candidates)")

    return unique


# ── PHASE 2: TRIAGE ───────────────────────────────────────────────────────────

TRIAGE_SYSTEM = """\
Eres un asistente de investigación para el Repositorio de Algoritmos Públicos de GobLab UAI (Chile).
Evalúa si un artículo describe un sistema de IA o algoritmo ESPECÍFICO que está siendo usado \
u oficialmente planificado por una institución pública chilena.

Responde ÚNICAMENTE con JSON válido con estas claves exactas:
{
  "relevant": true o false,
  "reason": "una oración explicando la decisión",
  "system_name": "nombre del sistema/proyecto, o null",
  "institution": "nombre de la institución pública chilena, o null"
}

relevant = true SOLO si se cumplen TODOS estos criterios:
- Describe un sistema, herramienta, o proyecto ESPECÍFICO (no tendencias generales de IA)
- Involucra al sector público chileno (gobierno, ministerios, servicios públicos, municipios)
- El sistema está activo, en piloto oficial, o en planificación formal
- El artículo tiene contenido textual sustancial (no es principalmente video o multimedia)

relevant = false si:
- Es sobre tendencias generales de IA sin sistema específico
- Involucra solo sector privado sin contrato o uso público
- Habla de otro país y no del sector público de Chile específicamente
- Es principalmente un video, ranking, opinión, o entretenimiento
- La mención de Chile o de IA es incidental o de relleno\
"""


def triage_one(candidate, client):
    url  = candidate["url"]
    text = fetch_url(url)

    # Pre-filter: skip failed fetches or suspiciously short content
    if len(text) < 300 or text.startswith("[ERROR"):
        return {
            "url":         url,
            "relevant":    False,
            "reason":      "Contenido insuficiente o error al obtener la página",
            "system_name": None,
            "institution": None,
            "orig_source": candidate.get("source", ""),
        }

    truncated = text[:4000]

    try:
        resp = client.messages.create(
            model=TRIAGE_MODEL,
            max_tokens=300,
            system=TRIAGE_SYSTEM,
            messages=[{"role": "user", "content": f"URL: {url}\n\nContenido:\n{truncated}"}],
        )
        raw = resp.content[0].text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        result = json.loads(raw)
    except Exception as e:
        result = {"relevant": False, "reason": f"Error en triage: {e}", "system_name": None, "institution": None}

    result["url"]         = url
    result["orig_source"] = candidate.get("source", "")
    return result


def phase_triage(candidates, client):
    print("\n" + "═" * 65)
    print("  FASE 2 — TRIAGE")
    print(f"  Claude ({TRIAGE_MODEL}) lee cada artículo y evalúa relevancia")
    print("═" * 65)
    print(f"\n  Analizando {len(candidates)} URLs... (esto puede tardar varios minutos)\n")

    results = [None] * len(candidates)
    done    = 0

    # 6 concurrent workers — polite to both sites and the API
    with ThreadPoolExecutor(max_workers=6) as pool:
        future_to_idx = {pool.submit(triage_one, c, client): i for i, c in enumerate(candidates)}
        for future in as_completed(future_to_idx):
            idx         = future_to_idx[future]
            result      = future.result()
            results[idx] = result
            done        += 1

            mark  = "✓" if result.get("relevant") else "✗"
            label = f"{result.get('system_name') or '—'} | {result.get('institution') or '—'}"
            short = result["url"][:65]
            print(f"  [{done:>3}/{len(candidates)}] {mark}  {short}")
            if result.get("relevant"):
                print(f"             → {label}")

    relevant = [r for r in results if r and r.get("relevant")]
    print(f"\n  → {len(relevant)} relevantes de {len(candidates)} analizadas")
    return relevant


# ── PHASE 3: GROUP ────────────────────────────────────────────────────────────

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


def phase_group(relevant, client):
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
        resp = client.messages.create(
            model=TRIAGE_MODEL,
            max_tokens=2000,
            system=GROUP_SYSTEM,
            messages=[{"role": "user", "content": f"Agrupa estos artículos:\n\n{lines}"}],
        )
        raw = resp.content[0].text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        parsed = json.loads(raw)

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
        print(f"  ERROR al agrupar: {e} — cada URL tratada como proyecto independiente")
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

def phase_generate(groups, client, output_path):
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
            blocks.append({"url": item["url"], "text": text})

        print(f"    Enviando {len(blocks)} fuente(s) a Claude para extracción...")
        ficha = extract_ficha_fields(blocks, client)
        fichas.append(ficha)
        source_map.append(blocks)

        if "_parse_error" not in ficha:
            n_sources = len(ficha.get("fuentes_apa") or [])
            print(f"    → {ficha.get('nombre', '?')} | {ficha.get('institucion_publica', '?')} | {n_sources} cita(s) APA")

    OUTPUT_DIR.mkdir(exist_ok=True)
    build_docx(fichas, source_map, output_path)


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="GobLab pipeline: scrape → triage → group → fichas"
    )
    parser.add_argument("--max-candidates", type=int, default=None,
                        help="Cap the number of URLs sent to triage (useful for quick tests)")
    parser.add_argument("--output", "-o", default=None,
                        help="Output .docx filename (saved in fichas_output/)")
    args = parser.parse_args()

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: Set ANTHROPIC_API_KEY environment variable.")
        sys.exit(1)

    client    = anthropic.Anthropic(api_key=api_key)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name  = args.output or f"pipeline_{timestamp}.docx"
    out_path  = OUTPUT_DIR / out_name

    print("\n" + "═" * 65)
    print("  GobLab — Pipeline automatizado")
    print(f"  Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"  Salida: fichas_output/{out_name}")
    print("═" * 65)

    t0 = time.time()

    candidates = phase_scrape(args.max_candidates)
    relevant   = phase_triage(candidates, client)
    groups     = phase_group(relevant, client)

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
