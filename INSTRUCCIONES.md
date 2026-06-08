# Generador de Fichas — Repositorio Algoritmos Públicos

Genera fichas resumidas para el Comité Editorial a partir de URLs o PDFs.

## Setup (una sola vez)

```bash
python3 -m pip install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."   # tu clave de API de Anthropic
```

## Uso

### 1 URL → 1 ficha
```bash
python3 generate_ficha.py https://www.ips.gob.cl/...
```

### Varias URLs del mismo algoritmo → 1 ficha (usar --group)
```bash
python3 generate_ficha.py https://fuente1.cl https://fuente2.cl --group
```

### Varias URLs de algoritmos distintos → N fichas
```bash
python3 generate_ficha.py https://algoritmo1.cl https://algoritmo2.cl
```

### Desde un archivo de texto (una URL por línea)
```bash
python3 generate_ficha.py --file mis_urls.txt
```

### Desde un PDF local
```bash
python3 generate_ficha.py --pdf documento.pdf
```

### Combinado + nombre de salida personalizado
```bash
python3 generate_ficha.py --file urls.txt --pdf informe.pdf --group --output fichas_julio.docx
```

## Output

Los archivos Word se guardan en la carpeta `fichas_output/`.

Cada ficha contiene:
- Nombre del sistema
- Objetivo del sistema
- Decisión que se automatiza
- Tipo (Automático / Semiautomático)
- Uso de datos personales (Sí/No)
- Institución Pública
- Unidad/Dirección
- Ejecutor
- Fuentes en formato APA
- Notas para el Comité
- Casilleros de aprobación (para marcar manualmente)

## Notas

- Si el sitio requiere JavaScript para cargar contenido, la herramienta puede capturar poco texto. En ese caso, copiar el texto manualmente a un .txt y usar `--pdf` o agregar el texto en un archivo.
- Los campos con información insuficiente aparecen como "Sin información disponible".
- La API de Claude hace la extracción — revisar siempre las fichas antes de presentar al Comité.
