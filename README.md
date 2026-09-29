# Consulta de roles de avalúo del SII

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) ![CLI](https://img.shields.io/badge/CLI-4D4D4D?style=flat-square&logo=gnubash&logoColor=white) ![GeoJSON](https://img.shields.io/badge/GeoJSON-3A7BD5?style=flat-square)

Herramienta de línea de comandos, en un solo archivo Python, para consultar de forma
**masiva** información de predios en la [Cartografía Digital del SII (Mapas SII)](https://www4.sii.cl/mapasui/internet/)
a partir de **comuna + manzana + predio** (el rol de avalúo).

Por cada predio obtiene su **comuna, dirección y coordenada geográfica**, y entrega los
resultados en cuatro formatos listos para usar: **CSV**, **Excel**, **GeoJSON** y una
**página HTML interactiva con mapa**.

Pensada para trabajo territorial y catastral: permite reemplazar la consulta manual —predio
por predio en el visor web— por un proceso por lotes reproducible.

> ℹ️ Usa información **pública y referencial** del SII, mediante consultas anónimas. No
> requiere inicio de sesión ni credenciales.

---

## Características

- **Entrada por lotes** desde un CSV (`COMUNA`, `MANZANA`, `PREDIO`) o consulta individual por argumentos.
- **Catálogo de las 347 comunas** embebido: acepta el nombre de la comuna o su código SII.
- **Georreferenciación**: entrega latitud/longitud (WGS84) de cada predio.
- **Cuatro salidas** desde una sola corrida, todas a una carpeta `outputs/`:
  - `CSV` — para pipelines y SIG.
  - `Excel (.xlsx)` — con las coordenadas como números reales (a prueba de la conversión regional de Excel).
  - `GeoJSON` — puntos listos para QGIS / ArcGIS.
  - `HTML` — visor con mapa (Leaflet), tabla plegable por comuna, filtro y enlace a Google Maps.
- **Robusto para lotes grandes**: pausa configurable entre consultas, reintentos con espera progresiva y manejo de codificaciones (UTF-8 / cp1252 / latin-1).
- **Sin dependencias pesadas**: solo `requests` (y `openpyxl` opcional para el Excel).

## Cómo funciona

El visor de Mapas SII consulta sus datos mediante un servicio JSON interno
(`mapasFacadeService/getPredioNacional`). Esta herramienta reproduce esa misma llamada
—descubierta inspeccionando el tráfico de red del visor— enviando `comuna/manzana/predio`
y quedándose con los campos de interés de la respuesta. Es, en esencia, un cliente pequeño
y respetuoso de un servicio público ya existente.

## Requisitos

- Python 3.8 o superior
- Dependencias:

```bash
pip install requests          # obligatorio
pip install openpyxl          # opcional, solo para el Excel
```

## Uso

### Lote desde CSV

Prepara un CSV con las columnas **COMUNA**, **MANZANA**, **PREDIO** (ver
[`plantilla_entrada.csv`](plantilla_entrada.csv)) y ejecuta:

```bash
python sii_roles.py entrada.csv
```

Los resultados quedan en la carpeta `outputs/` (`salida_sii.csv`, `.xlsx`, `.geojson`, `.html`).

> 📁 En [`outputs/`](outputs/) se incluyen **salidas de ejemplo** (`ejemplo_sii.*`) generadas
> a partir de [`plantilla_entrada.csv`](plantilla_entrada.csv), para ver los cuatro formatos
> sin necesidad de ejecutar la herramienta. Las corridas reales (`salida_sii.*`) quedan
> excluidas del repositorio por `.gitignore`.

### Consulta individual

```bash
python sii_roles.py --comuna RANCAGUA --manzana 1 --predio 1
# la comuna acepta nombre o código SII (--comuna 6101); alternativa: --rol 1-1
```

### Listar códigos de comuna

```bash
python sii_roles.py --list-comunas rancagua
```

## Opciones

| Opción | Descripción |
|---|---|
| `-o NOMBRE` | Nombre base de las salidas (por defecto `salida_sii`). |
| `--outdir CARPETA` | Carpeta de salida (por defecto `outputs`). |
| `--no-csv` / `--no-xlsx` / `--no-geojson` / `--no-html` | Omitir un formato. |
| `--delay 0.5` | Pausa en segundos entre consultas. |
| `--reintentos 3` | Reintentos ante errores de red. |
| `--col-comuna` / `--col-manzana` / `--col-predio` | Forzar el nombre de las columnas del CSV. |

## Formato de entrada

```csv
COMUNA;MANZANA;PREDIO
RANCAGUA;1;1
SAN FERNANDO;50;3
```

Acepta separador `,` o `;`, y comuna por nombre o por código SII.

## Nota sobre coordenadas y Excel

El CSV usa **punto decimal** (`-70.748958`), estándar en SIG. Al abrir ese CSV directamente
en Excel con configuración regional chilena, Excel puede interpretar el punto como separador
de miles y descuadrar el valor. Por eso la herramienta genera además un **`.xlsx`** con las
coordenadas ya como números reales. Recomendación: para revisar en Excel usa el `.xlsx`; para
SIG, el `.csv` o el `.geojson`.

## Aviso

Proyecto personal, sin afiliación ni respaldo del Servicio de Impuestos Internos. Consulta
información pública referencial; el propio SII indica que estos datos son referenciales y no
constituyen certificación oficial. Úsese de forma responsable y con pausas razonables entre
consultas.

## Licencia

[MIT](LICENSE) © Nelson Sánchez ([@leosanchez92](https://github.com/leosanchez92))
