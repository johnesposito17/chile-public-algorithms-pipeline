# Repositorio de Algoritmos Públicos — Pipeline Automatizado
### Automated Public Algorithms Repository Pipeline
#### GobLab UAI · Universidad Adolfo Ibáñez · Santiago, Chile

---

> **Ir a:** [Español](#español) · [English](#english)

---

## Español

### Descripción General

Este repositorio contiene el código de automatización desarrollado para apoyar la investigación y el mantenimiento del **[Repositorio de Algoritmos Públicos](https://algoritmos.gob.cl)** de **GobLab**, el laboratorio de gobierno de la Universidad Adolfo Ibáñez (UAI) en Santiago, Chile.

El Repositorio de Algoritmos Públicos es un registro nacional que cataloga los sistemas algorítmicos e inteligencia artificial utilizados por instituciones del sector público chileno. El objetivo es generar transparencia, promover la rendición de cuentas y facilitar el debate informado sobre el uso de IA en el Estado.

El pipeline automatizado en este repositorio **busca, filtra y documenta** nuevos sistemas de IA en el sector público que aún no han sido catalogados, reduciendo significativamente el tiempo de investigación manual del equipo editorial.

---

### Estructura del Repositorio

```
├── pipeline.py               Pipeline completo: scraping → triage → agrupación → fichas
├── generate_ficha.py         Generador individual de fichas (acepta URLs, PDFs, texto)
├── requirements.txt          Dependencias de Python
├── INSTRUCCIONES.md          Guía de uso en español
├── given materials/
│   └── Organización Casos Repositorio - Proyectos Repositorio.csv
│                             Base de datos actual del Repositorio (ver abajo)
└── examples/
    ├── fichas_muestra_jun2026.docx   Muestra de fichas generadas por el pipeline (jun 2026)
    ├── fichas_completas_jun2026.docx Ejecución completa con ~20 candidatos (jun 2026)
    └── notable_findings.md           5 hallazgos destacados con análisis de relevancia
```

---

### Base de Datos de Referencia

El archivo `given materials/Organización Casos Repositorio - Proyectos Repositorio.csv` contiene la versión más reciente de los **170 algoritmos** actualmente catalogados en el Repositorio. Este archivo fue exportado desde la planilla colaborativa del equipo editorial y corresponde a la versión actualizada tras la última reunión del comité de académicos y representantes de instituciones de investigación de distintos países de América del Sur.

El pipeline usa esta base de datos para dos propósitos:
1. **Deduplicación:** evitar proponer algoritmos que ya están en el Repositorio.
2. **Contexto para el triage:** Claude recibe el listado completo para distinguir sistemas nuevos de conocidos.

Para mantener la base de datos actualizada, reemplazar este archivo con la exportación CSV más reciente de la planilla antes de cada ejecución mensual.

---

### Novedades — Junio 2026

Las siguientes mejoras fueron implementadas en la versión actual del pipeline:

#### ✦ Filtrado de contenido propio de GobLab
El script ahora excluye explícitamente artículos sobre el propio Repositorio de Algoritmos Públicos de GobLab UAI. En versiones anteriores, menciones al Repositorio mismo podían pasar el filtro de relevancia, generando entradas redundantes.

#### ✦ Prompt Caching — reducción de costos ~10×
Se implementó el sistema de caché de prompts de la API de Anthropic (`cache_control: ephemeral`). El resumen de los 170 algoritmos existentes (~7.000 tokens) se envía con cada llamada de triage pero ahora se marca como cacheable. A partir de la segunda llamada en una ejecución, el costo de los tokens del sistema se reduce aproximadamente un 90%, manteniendo idéntica calidad y velocidad. En una ejecución típica de 500 URLs, esto equivale a un ahorro de ~USD 3–4 por ejecución.

#### ✦ Nueva fuente: MercadoPublico / ChileCompra
Se incorporó un módulo de scraping para **mercadopublico.cl** (el portal oficial de compras públicas de Chile) y el RSS de **chilecompra.cl**. Las licitaciones públicas formalizan el uso de sistemas algorítmicos meses antes de que aparezcan en medios de prensa, lo que convierte a MercadoPublico en una fuente de alta anticipación. El módulo usa Playwright para navegar el portal JavaScript-renderizado y extraer URLs de licitaciones relacionadas con IA.

#### ✦ Memoria de URLs (`seen_urls.json`)
Cada URL analizada se guarda en un archivo local. En las ejecuciones mensuales siguientes, estas URLs se omiten automáticamente, evitando reprocesar artículos ya evaluados y reduciendo el costo y tiempo de cada ejecución.

---

### Fuentes de Datos

| Fuente | Tipo | Descripción |
|--------|------|-------------|
| Google News RSS | 20 términos con filtro de fecha | Artículos del último año sobre IA en sector público chileno |
| Radio Universidad de Chile | RSS directo | Cobertura de políticas públicas y ciencia |
| La Tercera | RSS directo | Diario de referencia nacional |
| CIPER Chile | RSS directo | Periodismo investigativo |
| The Clinic | RSS directo | Periodismo político e investigativo |
| ANID | RSS + scraping | Agencia Nacional de Investigación y Desarrollo |
| ChileCompra | RSS directo | Noticias institucionales y licitaciones |
| MercadoPublico | Playwright (JS) | Portal de compras públicas — licitaciones directas |
| gob.cl | Scraping HTML | Portal de noticias del Gobierno de Chile |
| digital.gob.cl | Scraping HTML | División Gobierno Digital |
| Ministerio de Ciencia | Scraping HTML | Noticias de política científica |
| TransformacionPublica | Scraping HTML | Blog de modernización del Estado |
| Laboratorio de Gobierno | Scraping HTML | Soluciones e innovación pública |
| Ministerio de Hacienda | Scraping HTML | Noticias de hacienda |
| Ministerio de Economía | Scraping HTML | Noticias económicas |
| Servicio Civil | Scraping HTML | Gestión de personas en el Estado |
| DIPRES | Scraping HTML | Balances de Gestión Integral |

---

### Instalación y Uso

**Requisitos previos:** Python 3.11+, clave de API de Anthropic.

```bash
# Instalar dependencias
pip3 install -r requirements.txt
playwright install chromium

# Configurar clave de API
export ANTHROPIC_API_KEY="sk-ant-..."

# Ejecutar pipeline completo
python3 pipeline.py

# Ejecución de prueba (50 URLs)
python3 pipeline.py --max-candidates 50

# Nombre de archivo personalizado
python3 pipeline.py --output fichas_julio_2026.docx

# Generar ficha para una URL específica
python3 generate_ficha.py https://www.ejemplo.cl/noticia-sobre-algoritmo
```

Los archivos Word con las fichas se generan en `fichas_output/`.

---

### Costos Estimados (por ejecución mensual)

| Componente | Sin caché | Con caché (actual) |
|------------|-----------|---------------------|
| Triage (~500 URLs) | ~USD 3.50 | ~USD 0.50 |
| Generación de fichas (~10 proyectos) | ~USD 1.50 | ~USD 1.50 |
| **Total estimado** | **~USD 5.00** | **~USD 2.00** |

A partir de la segunda ejecución, el módulo de memoria de URLs reduce adicionalmente el costo al omitir URLs ya procesadas.

---

### Ejemplos de Salida

La carpeta `examples/` contiene:
- **`fichas_muestra_jun2026.docx`** — Ejecución del 11 de junio de 2026 con 20 candidatos identificados, incluyendo Latam-GPT, el Observatorio de Anomalías de ChileCompra, y AI-MarketScan.
- **`fichas_completas_jun2026.docx`** — Ejecución del 9 de junio de 2026 con ~48 candidatos evaluados, mayor cobertura.
- **`notable_findings.md`** — Análisis detallado de 5 hallazgos de alta relevancia identificados durante las pruebas.

---

### Notas Técnicas

- Las URLs con JavaScript puro (Google News redirects, MercadoPublico) se resuelven con Playwright (Chromium headless).
- El pipeline usa Claude Haiku para el triage masivo y Claude Opus para la extracción detallada de fichas.
- Los campos sin información suficiente aparecen como "Sin información disponible" — esto es preferible a datos incorrectos.
- Revisar siempre las fichas generadas antes de presentar al Comité Editorial.

---
---

## English

### Overview

This repository contains the automation code developed to support the research and maintenance of the **[Repositorio de Algoritmos Públicos](https://algoritmos.gob.cl)** (Public Algorithms Repository) at **GobLab**, the government laboratory of Universidad Adolfo Ibáñez (UAI) in Santiago, Chile.

The Public Algorithms Repository is a national registry cataloguing algorithmic and AI systems used by Chilean public sector institutions. Its goal is to promote transparency, accountability, and informed public debate about the use of AI in government.

The automated pipeline in this repository **finds, filters, and documents** new AI systems in the public sector that have not yet been catalogued, significantly reducing the research team's manual workload.

---

### Repository Structure

```
├── pipeline.py               Full pipeline: scraping → triage → grouping → fichas
├── generate_ficha.py         Individual ficha generator (accepts URLs, PDFs, text)
├── requirements.txt          Python dependencies
├── INSTRUCCIONES.md          Usage guide (Spanish)
├── given materials/
│   └── Organización Casos Repositorio - Proyectos Repositorio.csv
│                             Current Repository database (see below)
└── examples/
    ├── fichas_muestra_jun2026.docx   Sample fichas generated by the pipeline (Jun 2026)
    ├── fichas_completas_jun2026.docx Full run with ~20 candidates (Jun 2026)
    └── notable_findings.md           5 highlighted findings with relevance analysis
```

---

### Reference Database

The file `given materials/Organización Casos Repositorio - Proyectos Repositorio.csv` contains the most recent version of the **170 algorithms** currently catalogued in the Repository. This file was exported from the editorial team's collaborative spreadsheet and reflects the updated version from the most recent meeting of the committee of scholars and representatives from research institutions across South America.

The pipeline uses this database for two purposes:
1. **Deduplication:** to avoid proposing algorithms already in the Repository.
2. **Triage context:** Claude receives the full list to distinguish new systems from known ones.

To keep the database current, replace this file with the latest CSV export from the spreadsheet before each monthly run.

---

### What's New — June 2026

The following improvements were implemented in the current version:

#### ✦ Filtering of GobLab's Own Content
The script now explicitly excludes articles about GobLab UAI's own Public Algorithms Repository. In earlier versions, mentions of the Repository itself could pass the relevance filter, generating redundant entries.

#### ✦ Prompt Caching — ~10× Cost Reduction
Anthropic's prompt caching API (`cache_control: ephemeral`) was implemented. The summary of the 170 existing algorithms (~7,000 tokens) is sent with every triage call but is now marked as cacheable. From the second call onward within a run, the system token cost is reduced by approximately 90%, with identical quality and speed. On a typical 500-URL run, this saves ~USD 3–4 per execution.

#### ✦ New Source: MercadoPublico / ChileCompra
A scraping module was added for **mercadopublico.cl** (Chile's official public procurement portal) and the **chilecompra.cl** RSS feed. Public procurement tenders formalize the use of algorithmic systems months before they appear in press coverage, making MercadoPublico a high-anticipation source. The module uses Playwright to navigate the JavaScript-rendered portal and extract tender URLs related to AI.

#### ✦ URL Memory (`seen_urls.json`)
Every analyzed URL is saved to a local file. In subsequent monthly runs, these URLs are automatically skipped, avoiding reprocessing of already-evaluated articles and reducing the cost and time of each run.

---

### Data Sources

| Source | Type | Description |
|--------|------|-------------|
| Google News RSS | 20 terms with date filter | Articles from the past year on AI in Chilean public sector |
| Radio Universidad de Chile | Direct RSS | Public policy and science coverage |
| La Tercera | Direct RSS | National reference newspaper |
| CIPER Chile | Direct RSS | Investigative journalism |
| The Clinic | Direct RSS | Political and investigative journalism |
| ANID | RSS + scraping | National Research and Development Agency |
| ChileCompra | Direct RSS | Institutional news and tenders |
| MercadoPublico | Playwright (JS) | Public procurement portal — direct tenders |
| gob.cl | HTML scraping | Government of Chile news portal |
| digital.gob.cl | HTML scraping | Digital Government Division |
| Ministry of Science | HTML scraping | Science policy news |
| TransformacionPublica | HTML scraping | State modernization blog |
| Laboratorio de Gobierno | HTML scraping | Government Lab solutions |
| Ministry of Finance | HTML scraping | Finance news |
| Ministry of Economy | HTML scraping | Economic news |
| Servicio Civil | HTML scraping | Civil service management |
| DIPRES | HTML scraping | Integrated Management Reports |

---

### Installation and Usage

**Prerequisites:** Python 3.11+, Anthropic API key.

```bash
# Install dependencies
pip3 install -r requirements.txt
playwright install chromium

# Set API key
export ANTHROPIC_API_KEY="sk-ant-..."

# Run full pipeline
python3 pipeline.py

# Test run (50 URLs)
python3 pipeline.py --max-candidates 50

# Custom output filename
python3 pipeline.py --output fichas_july_2026.docx

# Generate ficha for a specific URL
python3 generate_ficha.py https://www.example.cl/algorithm-article
```

Word files with fichas are saved to `fichas_output/`.

---

### Estimated Costs (per monthly run)

| Component | Without cache | With cache (current) |
|-----------|--------------|----------------------|
| Triage (~500 URLs) | ~USD 3.50 | ~USD 0.50 |
| Ficha generation (~10 projects) | ~USD 1.50 | ~USD 1.50 |
| **Total estimate** | **~USD 5.00** | **~USD 2.00** |

From the second run onward, the URL memory module further reduces cost by skipping already-processed URLs.

---

### Example Output

The `examples/` folder contains:
- **`fichas_muestra_jun2026.docx`** — June 11, 2026 run with 20 identified candidates, including Latam-GPT, ChileCompra's Anomaly Observatory, and AI-MarketScan.
- **`fichas_completas_jun2026.docx`** — June 9, 2026 run with ~48 evaluated candidates, broader coverage.
- **`notable_findings.md`** — Detailed analysis of 5 high-relevance findings identified during testing.

---

### Technical Notes

- Pure JavaScript URLs (Google News redirects, MercadoPublico) are resolved with Playwright (headless Chromium).
- The pipeline uses Claude Haiku for mass triage and Claude Opus for detailed ficha extraction.
- Fields with insufficient information appear as "Sin información disponible" — this is preferable to incorrect data.
- Always review generated fichas before presenting to the Editorial Committee.

---

*Desarrollado por / Developed by: GobLab UAI — [goblab.uai.cl](https://goblab.uai.cl)*
