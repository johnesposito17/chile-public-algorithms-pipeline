# GobLab — Repositorio de Algoritmos Públicos

Herramientas para identificar y documentar sistemas de IA y algoritmos en el sector público chileno.

## Setup (una sola vez)

```bash
pip3 install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."
```

---

## Pipeline automatizado (recomendado)

Ejecuta todo de una vez: scraping → triage → agrupación → fichas.

```bash
python3 pipeline.py
```

### Opciones

```bash
# Limitar a 50 URLs para una prueba rápida
python3 pipeline.py --max-candidates 50

# Nombre de archivo de salida personalizado
python3 pipeline.py --output fichas_julio.docx
```

### Fases del pipeline

| Fase | Descripción | Modelo |
|------|-------------|--------|
| 1. Scrape | Google News RSS (8 términos) + sitios institucionales | — |
| 2. Triage | Claude lee cada artículo y descarta irrelevantes | claude-haiku |
| 3. Agrupación | Agrupa URLs que hablan del mismo proyecto | claude-haiku |
| 4. Fichas | Genera un .docx con una ficha por proyecto | claude-opus |

URLs de distintas fuentes sobre el mismo sistema se agrupan en **una sola ficha** con **múltiples citas APA**.

---

## Herramientas individuales

### Generar una ficha desde una URL específica

```bash
python3 generate_ficha.py https://www.ejemplo.cl/noticia
```

### Varias URLs del mismo algoritmo → 1 ficha

```bash
python3 generate_ficha.py https://fuente1.cl https://fuente2.cl --group
```

### Desde un archivo de texto (una URL por línea)

```bash
python3 generate_ficha.py --file mis_urls.txt
```

### Desde un PDF local

```bash
python3 generate_ficha.py --pdf documento.pdf
```

---

## Output

Los archivos Word se guardan en `fichas_output/`.

Cada ficha contiene:
- Nombre del sistema
- Objetivo del sistema
- Decisión que se automatiza
- Tipo (Automático / Semiautomático)
- Uso de datos personales (Sí/No)
- Institución Pública
- Unidad/Dirección
- Ejecutor
- Fuentes en formato APA (puede haber múltiples)
- Notas para el Comité
- Casilleros de aprobación

## Notas

- Si el sitio requiere JavaScript para cargar contenido, la herramienta puede capturar poco texto. En ese caso, copiar el texto manualmente y usar `--pdf`.
- Los campos con información insuficiente aparecen como "Sin información disponible".
- Revisar siempre las fichas antes de presentar al Comité.
