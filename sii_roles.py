#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sii_roles.py - Consulta masiva de roles de avaluo en el mapa del SII (Cartografia Digital SII).

Entrada: comuna (nombre o codigo SII) + manzana + predio.
Salida : comuna, direccion y coordenada del predio (lat/lon WGS84).
         Genera CSV, Excel (.xlsx, a prueba de Excel), GeoJSON de puntos y un HTML
         con mapa + tabla abatible. Todo se guarda en la carpeta outputs.

Fuente: POST https://www4.sii.cl/mapasui/services/data/mapasFacadeService/getPredioNacional
        (consulta anonima, informacion publica referencial del SII).

Uso rapido:
    # Una consulta
    python sii_roles.py --comuna RANCAGUA --manzana 1 --predio 1

    # Lote desde CSV (columnas: COMUNA, MANZANA, PREDIO) -> outputs\\...
    python sii_roles.py entrada.csv

    # Listar / buscar codigos de comuna
    python sii_roles.py --list-comunas RANCAGUA

Requiere:  Python 3.8+  y  el paquete  requests   (pip install requests)
           opcional: openpyxl para el Excel  (pip install openpyxl)
Autor: Nelson Sanchez  -  github.com/leosanchez92

Nota sobre coordenadas: el .csv guarda lat/lon con punto decimal (formato estandar GIS).
Si lo abres en Excel en configuracion regional chilena, Excel puede leer el punto como
separador de miles y descuadrar el valor. Por eso el .xlsx ya trae las coordenadas como
numeros reales (no se descuadran) y el .html/.geojson tampoco tienen ese problema.
Para GIS usa el .csv o el .geojson; para mirar en Excel usa el .xlsx.
"""

import argparse
import csv
import io
import json
import os
import sys
import time
import unicodedata
import uuid

try:
    import requests
except ImportError:
    sys.exit("Falta el paquete 'requests'. Instalalo con:  pip install requests")

# ---------------------------------------------------------------------------
# Configuracion del servicio
# ---------------------------------------------------------------------------
URL = "https://www4.sii.cl/mapasui/services/data/mapasFacadeService/getPredioNacional"
NAMESPACE = "cl.sii.sdi.lob.bbrr.mapas.data.api.interfaces.MapasFacadeService/getPredioNacional"

HEADERS = {
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://www4.sii.cl",
    "Referer": "https://www4.sii.cl/mapasui/internet/",
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
}

DIR_SALIDAS = "outputs"  # carpeta por defecto para todas las salidas

# Campos que se exportan (clave_api -> encabezado). Solo lo pedido: comuna, direccion, coordenada.
CAMPOS = [
    ("nombreComuna", "comuna"),
    ("direccion",    "direccion"),
    ("ubicacionY",   "lon"),
    ("ubicacionX",   "lat"),
]

# ---------------------------------------------------------------------------
# Catalogo de comunas (codigo SII <-> nombre). Embebido, sin dependencias.
# ---------------------------------------------------------------------------
_COMUNAS = json.loads(r"""[{"codigo":"5406","nombre":"ALGARROBO"},{"codigo":"14605","nombre":"ALHUÉ"},{"codigo":"8414","nombre":"ALTO BIOBÍO"},{"codigo":"3304","nombre":"ALTO DEL CARMEN"},{"codigo":"1211","nombre":"ALTO HOSPICIO"},{"codigo":"10406","nombre":"ANCUD"},{"codigo":"4104","nombre":"ANDACOLLO"},{"codigo":"9101","nombre":"ANGOL"},{"codigo":"2201","nombre":"ANTOFAGASTA"},{"codigo":"8413","nombre":"ANTUCO"},{"codigo":"8301","nombre":"ARAUCO"},{"codigo":"1101","nombre":"ARICA"},{"codigo":"11101","nombre":"AYSÉN"},{"codigo":"16403","nombre":"BUIN"},{"codigo":"8113","nombre":"BULNES"},{"codigo":"5203","nombre":"CABILDO"},{"codigo":"12401","nombre":"CABO DE HORNOS"},{"codigo":"8410","nombre":"CABRERO"},{"codigo":"2301","nombre":"CALAMA"},{"codigo":"10309","nombre":"CALBUCO"},{"codigo":"3202","nombre":"CALDERA"},{"codigo":"16402","nombre":"CALERA DE TANGO"},{"codigo":"5702","nombre":"CALLE LARGA"},{"codigo":"1106","nombre":"CAMARONES"},{"codigo":"1208","nombre":"CAMIÑA"},{"codigo":"4304","nombre":"CANELA"},{"codigo":"9209","nombre":"CARAHUE"},{"codigo":"5403","nombre":"CARTAGENA"},{"codigo":"5305","nombre":"CASABLANCA"},{"codigo":"10401","nombre":"CASTRO"},{"codigo":"5603","nombre":"CATEMU"},{"codigo":"7401","nombre":"CAUQUENES"},{"codigo":"8305","nombre":"CAÑETE"},{"codigo":"14166","nombre":"CERRILLOS"},{"codigo":"14156","nombre":"CERRO NAVIA"},{"codigo":"10501","nombre":"CHAITÉN"},{"codigo":"7403","nombre":"CHANCO"},{"codigo":"3101","nombre":"CHAÑARAL"},{"codigo":"8211","nombre":"CHIGUAYANTE"},{"codigo":"11201","nombre":"CHILE CHICO"},{"codigo":"8101","nombre":"CHILLÁN"},{"codigo":"8121","nombre":"CHILLÁN VIEJO"},{"codigo":"6202","nombre":"CHIMBARONGO"},{"codigo":"9221","nombre":"CHOLCHOL"},{"codigo":"10402","nombre":"CHONCHI"},{"codigo":"6209","nombre":"CHÉPICA"},{"codigo":"11102","nombre":"CISNES"},{"codigo":"8107","nombre":"COBQUECURA"},{"codigo":"10302","nombre":"COCHAMÓ"},{"codigo":"11301","nombre":"COCHRANE"},{"codigo":"6107","nombre":"CODEGUA"},{"codigo":"8120","nombre":"COELEMU"},{"codigo":"8103","nombre":"COIHUECO"},{"codigo":"6116","nombre":"COINCO"},{"codigo":"7303","nombre":"COLBÚN"},{"codigo":"1210","nombre":"COLCHANE"},{"codigo":"14201","nombre":"COLINA"},{"codigo":"9105","nombre":"COLLIPULLI"},{"codigo":"6106","nombre":"COLTAUCO"},{"codigo":"4205","nombre":"COMBARBALÁ"},{"codigo":"8201","nombre":"CONCEPCIÓN"},{"codigo":"14127","nombre":"CONCHALÍ"},{"codigo":"5309","nombre":"CONCÓN"},{"codigo":"7208","nombre":"CONSTITUCIÓN"},{"codigo":"8306","nombre":"CONTULMO"},{"codigo":"3201","nombre":"COPIAPÓ"},{"codigo":"4103","nombre":"COQUIMBO"},{"codigo":"8207","nombre":"CORONEL"},{"codigo":"10106","nombre":"CORRAL"},{"codigo":"11401","nombre":"COYHAIQUE"},{"codigo":"9204","nombre":"CUNCO"},{"codigo":"9110","nombre":"CURACAUTÍN"},{"codigo":"14603","nombre":"CURACAVÍ"},{"codigo":"10410","nombre":"CURACO DE VÉLEZ"},{"codigo":"8302","nombre":"CURANILAHUE"},{"codigo":"9218","nombre":"CURARREHUE"},{"codigo":"7207","nombre":"CUREPTO"},{"codigo":"7101","nombre":"CURICÓ"},{"codigo":"10408","nombre":"DALCAHUE"},{"codigo":"3102","nombre":"DIEGO DE ALMAGRO"},{"codigo":"6105","nombre":"DOÑIHUE"},{"codigo":"16165","nombre":"EL BOSQUE"},{"codigo":"8118","nombre":"EL CARMEN"},{"codigo":"14503","nombre":"EL MONTE"},{"codigo":"5405","nombre":"EL QUISCO"},{"codigo":"5404","nombre":"EL TABO"},{"codigo":"7209","nombre":"EMPEDRADO"},{"codigo":"9106","nombre":"ERCILLA"},{"codigo":"14157","nombre":"ESTACIÓN CENTRAL"},{"codigo":"8204","nombre":"FLORIDA"},{"codigo":"9203","nombre":"FREIRE"},{"codigo":"3302","nombre":"FREIRINA"},{"codigo":"10304","nombre":"FRESIA"},{"codigo":"10305","nombre":"FRUTILLAR"},{"codigo":"10503","nombre":"FUTALEUFÚ"},{"codigo":"10105","nombre":"FUTRONO"},{"codigo":"9207","nombre":"GALVARINO"},{"codigo":"1302","nombre":"GENERAL LAGOS"},{"codigo":"9212","nombre":"GORBEA"},{"codigo":"6103","nombre":"GRANEROS"},{"codigo":"11104","nombre":"GUAITECAS"},{"codigo":"5503","nombre":"HIJUELAS"},{"codigo":"10502","nombre":"HUALAIHUÉ"},{"codigo":"7107","nombre":"HUALAÑÉ"},{"codigo":"8212","nombre":"HUALPÉN"},{"codigo":"8203","nombre":"HUALQUI"},{"codigo":"1206","nombre":"HUARA"},{"codigo":"3303","nombre":"HUASCO"},{"codigo":"14158","nombre":"HUECHURABA"},{"codigo":"4301","nombre":"ILLAPEL"},{"codigo":"13167","nombre":"INDEPENDENCIA"},{"codigo":"1201","nombre":"IQUIQUE"},{"codigo":"14502","nombre":"ISLA DE MAIPO"},{"codigo":"5101","nombre":"ISLA DE PASCUA"},{"codigo":"5308","nombre":"JUAN FERNÁNDEZ"},{"codigo":"5504","nombre":"LA CALERA"},{"codigo":"16110","nombre":"LA CISTERNA"},{"codigo":"5505","nombre":"LA CRUZ"},{"codigo":"6304","nombre":"LA ESTRELLA"},{"codigo":"15128","nombre":"LA FLORIDA"},{"codigo":"16131","nombre":"LA GRANJA"},{"codigo":"4102","nombre":"LA HIGUERA"},{"codigo":"5201","nombre":"LA LIGUA"},{"codigo":"16154","nombre":"LA PINTANA"},{"codigo":"15132","nombre":"LA REINA"},{"codigo":"4101","nombre":"LA SERENA"},{"codigo":"10109","nombre":"LA UNIÓN"},{"codigo":"10112","nombre":"LAGO RANCO"},{"codigo":"11402","nombre":"LAGO VERDE"},{"codigo":"12206","nombre":"LAGUNA BLANCA"},{"codigo":"8403","nombre":"LAJA"},{"codigo":"14202","nombre":"LAMPA"},{"codigo":"10103","nombre":"LANCO"},{"codigo":"6109","nombre":"LAS CABRAS"},{"codigo":"15108","nombre":"LAS CONDES"},{"codigo":"9205","nombre":"LAUTARO"},{"codigo":"8303","nombre":"LEBU"},{"codigo":"7105","nombre":"LICANTÉN"},{"codigo":"5506","nombre":"LIMACHE"},{"codigo":"7301","nombre":"LINARES"},{"codigo":"6303","nombre":"LITUECHE"},{"codigo":"10306","nombre":"LLANQUIHUE"},{"codigo":"5606","nombre":"LLAY-LLAY"},{"codigo":"15161","nombre":"LO BARNECHEA"},{"codigo":"16164","nombre":"LO ESPEJO"},{"codigo":"14155","nombre":"LO PRADO"},{"codigo":"6206","nombre":"LOLOL"},{"codigo":"9214","nombre":"LONCOCHE"},{"codigo":"7304","nombre":"LONGAVÍ"},{"codigo":"9111","nombre":"LONQUIMAY"},{"codigo":"8304","nombre":"LOS ALAMOS"},{"codigo":"5701","nombre":"LOS ANDES"},{"codigo":"8401","nombre":"LOS ANGELES"},{"codigo":"10104","nombre":"LOS LAGOS"},{"codigo":"10308","nombre":"LOS MUERMOS"},{"codigo":"9103","nombre":"LOS SAUCES"},{"codigo":"4303","nombre":"LOS VILOS"},{"codigo":"8208","nombre":"LOTA"},{"codigo":"9108","nombre":"LUMACO"},{"codigo":"6102","nombre":"MACHALÍ"},{"codigo":"15151","nombre":"MACUL"},{"codigo":"14109","nombre":"MAIPÚ"},{"codigo":"6115","nombre":"MALLOA"},{"codigo":"6305","nombre":"MARCHIGÜE"},{"codigo":"10102","nombre":"MARIQUINA"},{"codigo":"2103","nombre":"MARÍA ELENA"},{"codigo":"14602","nombre":"MARÍA PINTO"},{"codigo":"7206","nombre":"MAULE"},{"codigo":"10307","nombre":"MAULLÍN"},{"codigo":"2203","nombre":"MEJILLONES"},{"codigo":"9217","nombre":"MELIPEUCO"},{"codigo":"14601","nombre":"MELIPILLA"},{"codigo":"7108","nombre":"MOLINA"},{"codigo":"4203","nombre":"MONTE PATRIA"},{"codigo":"8407","nombre":"MULCHÉN"},{"codigo":"10107","nombre":"MÁFIL"},{"codigo":"8405","nombre":"NACIMIENTO"},{"codigo":"6203","nombre":"NANCAGUA"},{"codigo":"12101","nombre":"NATALES"},{"codigo":"6302","nombre":"NAVIDAD"},{"codigo":"8406","nombre":"NEGRETE"},{"codigo":"8105","nombre":"NINHUE"},{"codigo":"5502","nombre":"NOGALES"},{"codigo":"9208","nombre":"NUEVA IMPERIAL"},{"codigo":"11302","nombre":"O'HIGGINS"},{"codigo":"6114","nombre":"OLIVAR"},{"codigo":"2302","nombre":"OLLAGÜE"},{"codigo":"5507","nombre":"OLMUÉ"},{"codigo":"10201","nombre":"OSORNO"},{"codigo":"4201","nombre":"OVALLE"},{"codigo":"14505","nombre":"PADRE HURTADO"},{"codigo":"9220","nombre":"PADRE LAS CASAS"},{"codigo":"4106","nombre":"PAIHUANO"},{"codigo":"10110","nombre":"PAILLACO"},{"codigo":"16404","nombre":"PAINE"},{"codigo":"10504","nombre":"PALENA"},{"codigo":"6207","nombre":"PALMILLA"},{"codigo":"10108","nombre":"PANGUIPULLI"},{"codigo":"5602","nombre":"PANQUEHUE"},{"codigo":"5205","nombre":"PAPUDO"},{"codigo":"6306","nombre":"PAREDONES"},{"codigo":"7305","nombre":"PARRAL"},{"codigo":"16162","nombre":"PEDRO AGUIRRE CERDA"},{"codigo":"7203","nombre":"PELARCO"},{"codigo":"7402","nombre":"PELLUHUE"},{"codigo":"8117","nombre":"PEMUCO"},{"codigo":"7205","nombre":"PENCAHUE"},{"codigo":"8202","nombre":"PENCO"},{"codigo":"6208","nombre":"PERALILLO"},{"codigo":"9206","nombre":"PERQUENCO"},{"codigo":"5202","nombre":"PETORCA"},{"codigo":"6108","nombre":"PEUMO"},{"codigo":"14504","nombre":"PEÑAFLOR"},{"codigo":"15152","nombre":"PEÑALOLÉN"},{"codigo":"1203","nombre":"PICA"},{"codigo":"6111","nombre":"PICHIDEGUA"},{"codigo":"6301","nombre":"PICHILEMU"},{"codigo":"8102","nombre":"PINTO"},{"codigo":"16302","nombre":"PIRQUE"},{"codigo":"9211","nombre":"PITRUFQUÉN"},{"codigo":"6204","nombre":"PLACILLA"},{"codigo":"8106","nombre":"PORTEZUELO"},{"codigo":"12301","nombre":"PORVENIR"},{"codigo":"1204","nombre":"POZO ALMONTE"},{"codigo":"12302","nombre":"PRIMAVERA"},{"codigo":"15103","nombre":"PROVIDENCIA"},{"codigo":"5307","nombre":"PUCHUNCAVÍ"},{"codigo":"9216","nombre":"PUCÓN"},{"codigo":"14111","nombre":"PUDAHUEL"},{"codigo":"16301","nombre":"PUENTE ALTO"},{"codigo":"10301","nombre":"PUERTO MONTT"},{"codigo":"10203","nombre":"PUERTO OCTAY"},{"codigo":"10303","nombre":"PUERTO VARAS"},{"codigo":"6214","nombre":"PUMANQUE"},{"codigo":"4204","nombre":"PUNITAQUI"},{"codigo":"12205","nombre":"PUNTA ARENAS"},{"codigo":"10405","nombre":"PUQUELDÓN"},{"codigo":"10206","nombre":"PURRANQUE"},{"codigo":"9102","nombre":"PURÉN"},{"codigo":"5604","nombre":"PUTAENDO"},{"codigo":"1301","nombre":"PUTRE"},{"codigo":"10204","nombre":"PUYEHUE"},{"codigo":"10403","nombre":"QUEILÉN"},{"codigo":"10404","nombre":"QUELLÓN"},{"codigo":"10407","nombre":"QUEMCHI"},{"codigo":"8408","nombre":"QUILACO"},{"codigo":"14114","nombre":"QUILICURA"},{"codigo":"8404","nombre":"QUILLECO"},{"codigo":"5501","nombre":"QUILLOTA"},{"codigo":"8115","nombre":"QUILLÓN"},{"codigo":"5304","nombre":"QUILPUÉ"},{"codigo":"10415","nombre":"QUINCHAO"},{"codigo":"6117","nombre":"QUINTA DE TILCOCO"},{"codigo":"14107","nombre":"QUINTA NORMAL"},{"codigo":"5306","nombre":"QUINTERO"},{"codigo":"8104","nombre":"QUIRIHUE"},{"codigo":"6101","nombre":"RANCAGUA"},{"codigo":"7104","nombre":"RAUCO"},{"codigo":"13159","nombre":"RECOLETA"},{"codigo":"9104","nombre":"RENAICO"},{"codigo":"14113","nombre":"RENCA"},{"codigo":"6112","nombre":"RENGO"},{"codigo":"6113","nombre":"REQUÍNOA"},{"codigo":"7306","nombre":"RETIRO"},{"codigo":"5704","nombre":"RINCONADA"},{"codigo":"7103","nombre":"ROMERAL"},{"codigo":"8119","nombre":"RÁNQUIL"},{"codigo":"10111","nombre":"RÍO BUENO"},{"codigo":"7204","nombre":"RÍO CLARO"},{"codigo":"4206","nombre":"RÍO HURTADO"},{"codigo":"11203","nombre":"RÍO IBÁÑEZ"},{"codigo":"10205","nombre":"RÍO NEGRO"},{"codigo":"12202","nombre":"RÍO VERDE"},{"codigo":"9210","nombre":"SAAVEDRA"},{"codigo":"7109","nombre":"SAGRADA FAMILIA"},{"codigo":"4302","nombre":"SALAMANCA"},{"codigo":"5401","nombre":"SAN ANTONIO"},{"codigo":"16401","nombre":"SAN BERNARDO"},{"codigo":"8109","nombre":"SAN CARLOS"},{"codigo":"7202","nombre":"SAN CLEMENTE"},{"codigo":"5703","nombre":"SAN ESTEBAN"},{"codigo":"8111","nombre":"SAN FABIÁN"},{"codigo":"5601","nombre":"SAN FELIPE"},{"codigo":"6201","nombre":"SAN FERNANDO"},{"codigo":"6104","nombre":"SAN FRANCISCO DE MOSTAZAL"},{"codigo":"12204","nombre":"SAN GREGORIO"},{"codigo":"8114","nombre":"SAN IGNACIO"},{"codigo":"7310","nombre":"SAN JAVIER"},{"codigo":"16163","nombre":"SAN JOAQUÍN"},{"codigo":"16303","nombre":"SAN JOSÉ DE MAIPO"},{"codigo":"10207","nombre":"SAN JUAN DE LA COSTA"},{"codigo":"16106","nombre":"SAN MIGUEL"},{"codigo":"8112","nombre":"SAN NICOLÁS"},{"codigo":"10202","nombre":"SAN PABLO"},{"codigo":"14604","nombre":"SAN PEDRO"},{"codigo":"2303","nombre":"SAN PEDRO DE ATACAMA"},{"codigo":"8210","nombre":"SAN PEDRO DE LA PAZ"},{"codigo":"7210","nombre":"SAN RAFAEL"},{"codigo":"16153","nombre":"SAN RAMÓN"},{"codigo":"8411","nombre":"SAN ROSENDO"},{"codigo":"6110","nombre":"SAN VICENTE"},{"codigo":"8402","nombre":"SANTA BÁRBARA"},{"codigo":"6205","nombre":"SANTA CRUZ"},{"codigo":"8209","nombre":"SANTA JUANA"},{"codigo":"5605","nombre":"SANTA MARÍA"},{"codigo":"13101","nombre":"SANTIAGO"},{"codigo":"13134","nombre":"SANTIAGO OESTE"},{"codigo":"13135","nombre":"SANTIAGO SUR"},{"codigo":"5402","nombre":"SANTO DOMINGO"},{"codigo":"2206","nombre":"SIERRA GORDA"},{"codigo":"14501","nombre":"TALAGANTE"},{"codigo":"7201","nombre":"TALCA"},{"codigo":"8206","nombre":"TALCAHUANO"},{"codigo":"2202","nombre":"TALTAL"},{"codigo":"9201","nombre":"TEMUCO"},{"codigo":"7102","nombre":"TENO"},{"codigo":"9219","nombre":"TEODORO SCHMIDT"},{"codigo":"3203","nombre":"TIERRA AMARILLA"},{"codigo":"14203","nombre":"TIL-TIL"},{"codigo":"12304","nombre":"TIMAUKEL"},{"codigo":"8307","nombre":"TIRÚA"},{"codigo":"2101","nombre":"TOCOPILLA"},{"codigo":"9213","nombre":"TOLTÉN"},{"codigo":"8205","nombre":"TOMÉ"},{"codigo":"12103","nombre":"TORRES DEL PAINE"},{"codigo":"11303","nombre":"TORTEL"},{"codigo":"9107","nombre":"TRAIGUÉN"},{"codigo":"8108","nombre":"TREHUACO"},{"codigo":"8412","nombre":"TUCAPEL"},{"codigo":"10101","nombre":"VALDIVIA"},{"codigo":"3301","nombre":"VALLENAR"},{"codigo":"5301","nombre":"VALPARAISO"},{"codigo":"7106","nombre":"VICHUQUÉN"},{"codigo":"9109","nombre":"VICTORIA"},{"codigo":"4105","nombre":"VICUÑA"},{"codigo":"9202","nombre":"VILCÚN"},{"codigo":"7309","nombre":"VILLA ALEGRE"},{"codigo":"5303","nombre":"VILLA ALEMANA"},{"codigo":"9215","nombre":"VILLARRICA"},{"codigo":"15160","nombre":"VITACURA"},{"codigo":"5302","nombre":"VIÑA DEL MAR"},{"codigo":"7302","nombre":"YERBAS BUENAS"},{"codigo":"8409","nombre":"YUMBEL"},{"codigo":"8116","nombre":"YUNGAY"},{"codigo":"5204","nombre":"ZAPALLAR"},{"codigo":"8110","nombre":"ÑIQUÉN"},{"codigo":"15105","nombre":"ÑUÑOA"}]""")


def _norm(txt):
    if txt is None:
        return ""
    txt = unicodedata.normalize("NFD", str(txt))
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    return " ".join(txt.upper().replace("-", " ").replace("_", " ").split())


_POR_NOMBRE = {_norm(c["nombre"]): c["codigo"] for c in _COMUNAS}
_POR_CODIGO = {c["codigo"]: c["nombre"] for c in _COMUNAS}


def resolver_comuna(valor):
    """Devuelve (codigo, nombre). Acepta codigo SII o nombre de comuna."""
    v = str(valor).strip()
    if v in _POR_CODIGO:
        return v, _POR_CODIGO[v]
    n = _norm(v)
    if n in _POR_NOMBRE:
        cod = _POR_NOMBRE[n]
        return cod, _POR_CODIGO[cod]
    candidatos = [c for k, c in _POR_NOMBRE.items() if n and n in k]
    if len(set(candidatos)) == 1:
        cod = candidatos[0]
        return cod, _POR_CODIGO[cod]
    raise ValueError("Comuna no reconocida: %r (usa --list-comunas para ver opciones)" % valor)


def _entero(valor, campo):
    s = str(valor).strip()
    if not s.isdigit():
        raise ValueError("%s invalido: %r (debe ser numerico)" % (campo, valor))
    return int(s)


# ---------------------------------------------------------------------------
# Consulta al servicio
# ---------------------------------------------------------------------------
def consultar(session, comuna_cod, manzana, predio, reintentos=3, timeout=30):
    """Devuelve el dict 'data' del predio, o None si no existe."""
    body = {
        "metaData": {
            "namespace": NAMESPACE,
            "conversationId": "UNAUTHENTICATED-CALL",
            "transactionId": str(uuid.uuid4()),
        },
        "data": {
            "predio": {"comuna": str(comuna_cod), "manzana": int(manzana), "predio": int(predio)},
            "servicios": [],
            "tokenARSII": None,
        },
    }
    ultimo_error = None
    for intento in range(1, reintentos + 1):
        try:
            r = session.post(URL, json=body, headers=HEADERS, timeout=timeout)
            ctype = r.headers.get("Content-Type", "")
            if r.status_code == 200 and "application/json" in ctype:
                j = r.json()
                errores = (j.get("metaData") or {}).get("errors")
                if errores:
                    raise RuntimeError("El servicio devolvio errores: %s" % errores)
                return j.get("data")
            raise RuntimeError("Respuesta inesperada HTTP %s (%s)" % (r.status_code, ctype or "sin tipo"))
        except Exception as e:  # noqa: BLE001
            ultimo_error = e
            if intento < reintentos:
                time.sleep(min(2 ** intento, 8))
    raise RuntimeError(str(ultimo_error))


def fila_resultado(comuna_in, manzana_in, predio_in, comuna_nom, data, estado, detalle=""):
    fila = {
        "consulta_comuna": comuna_in,
        "consulta_manzana": manzana_in,
        "consulta_predio": predio_in,
        "estado": estado,
        "detalle": detalle,
    }
    for clave, encab in CAMPOS:
        val = "" if data is None else data.get(clave, "")
        fila[encab] = val.strip() if isinstance(val, str) else val
    if data is None and comuna_nom:
        fila["comuna"] = comuna_nom
    return fila


CAMPOS_SALIDA = (["consulta_comuna", "consulta_manzana", "consulta_predio", "estado", "detalle"]
                 + [encab for _, encab in CAMPOS])


# ---------------------------------------------------------------------------
# Lectura de la entrada (CSV con columnas COMUNA, MANZANA, PREDIO)
# ---------------------------------------------------------------------------
def _leer_texto(ruta):
    """Lee el CSV probando codificaciones habituales (Excel/Windows suele usar cp1252)."""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with open(ruta, "r", newline="", encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(ruta, "r", newline="", encoding="latin-1") as f:
        return f.read()


def leer_entrada(ruta, col_comuna=None, col_manzana=None, col_predio=None):
    """Devuelve lista de (comuna, manzana, predio) como strings."""
    texto = _leer_texto(ruta)
    try:
        dialecto = csv.Sniffer().sniff(texto[:4096], delimiters=",;\t")
    except csv.Error:
        dialecto = csv.excel
    filas_raw = [r for r in csv.reader(io.StringIO(texto), dialecto) if any(str(c).strip() for c in r)]
    if not filas_raw:
        return []

    encab = [_norm(c) for c in filas_raw[0]]

    def idx(forzado, alias):
        if forzado:
            fn = _norm(forzado)
            if fn in encab:
                return encab.index(fn)
        for a in alias:
            if a in encab:
                return encab.index(a)
        return None

    i_com = idx(col_comuna,  ["COMUNA", "COMUNA COD", "COD COMUNA", "CODIGO COMUNA"])
    i_man = idx(col_manzana, ["MANZANA", "MANZANA ROL", "ROL MANZANA", "MZ", "MANZ"])
    i_pre = idx(col_predio,  ["PREDIO", "PREDIO ROL", "ROL PREDIO", "PR", "PRED"])

    filas = []
    if i_com is not None and i_man is not None and i_pre is not None:
        for r in filas_raw[1:]:
            if len(r) <= max(i_com, i_man, i_pre):
                continue
            com, man, pre = r[i_com].strip(), r[i_man].strip(), r[i_pre].strip()
            if com or man or pre:
                filas.append((com, man, pre))
    else:
        for r in filas_raw:
            if len(r) >= 3 and any(c.strip() for c in r[:3]):
                filas.append((r[0].strip(), r[1].strip(), r[2].strip()))
    return filas


# ---------------------------------------------------------------------------
# Salidas: CSV, XLSX, GeoJSON, HTML
# ---------------------------------------------------------------------------
def escribir_csv(ruta, filas):
    with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS_SALIDA, extrasaction="ignore")
        w.writeheader()
        for fila in filas:
            w.writerow(fila)


def escribir_xlsx(ruta, filas):
    """Excel con lat/lon como NUMEROS reales -> Excel no los descuadra. Requiere openpyxl."""
    try:
        from openpyxl import Workbook
    except ImportError:
        return False
    wb = Workbook()
    ws = wb.active
    ws.title = "roles"
    ws.append(CAMPOS_SALIDA)
    for fila in filas:
        ws.append([fila.get(c, "") for c in CAMPOS_SALIDA])
    # formato de coordenadas con 6 decimales, punto decimal (no afecta el valor)
    col_lon = CAMPOS_SALIDA.index("lon") + 1
    col_lat = CAMPOS_SALIDA.index("lat") + 1
    for row in ws.iter_rows(min_row=2, min_col=min(col_lon, col_lat), max_col=max(col_lon, col_lat)):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = "0.000000"
    ws.freeze_panes = "A2"
    wb.save(ruta)
    return True


def escribir_geojson(ruta, filas):
    feats = []
    for fila in filas:
        try:
            lon = float(fila.get("lon"))
            lat = float(fila.get("lat"))
        except (TypeError, ValueError):
            continue
        props = {k: v for k, v in fila.items() if k not in ("lon", "lat")}
        feats.append({"type": "Feature",
                      "geometry": {"type": "Point", "coordinates": [lon, lat]},
                      "properties": props})
    fc = {"type": "FeatureCollection",
          "crs": {"type": "name", "properties": {"name": "urn:ogc:def:crs:OGC:1.3:CRS84"}},
          "features": feats}
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(fc, f, ensure_ascii=False, indent=1)
    return len(feats)


HTML_TEMPLATE = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Consulta de roles SII</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css">
<style>
  :root{--fg:#1c2733;--muted:#5b6b7a;--line:#e2e8f0;--bg:#f4f6f8;--card:#fff;--ok:#0f7b57;--no:#b45309;--accent:#2563eb}
  *{box-sizing:border-box}
  body{margin:0;font:14px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:var(--fg);background:var(--bg)}
  header{padding:14px 18px;background:var(--card);border-bottom:1px solid var(--line);position:sticky;top:0;z-index:1000}
  header h1{margin:0;font-size:17px}
  header p{margin:3px 0 0;color:var(--muted);font-size:12px}
  .stats{display:flex;gap:7px;flex-wrap:wrap;margin-top:9px}
  .chip{padding:3px 10px;border-radius:999px;font-size:12px;font-weight:600;background:var(--bg);border:1px solid var(--line)}
  .chip.ok{color:var(--ok)}.chip.no{color:var(--no)}
  #map{height:50vh;min-height:300px;width:100%;background:#dbe4ec}
  .wrap{padding:14px 18px 40px;max-width:1100px;margin:0 auto}
  .toolbar{display:flex;gap:10px;align-items:center;margin:4px 0 14px;flex-wrap:wrap}
  input[type=search]{padding:8px 11px;border:1px solid var(--line);border-radius:8px;font-size:13px;min-width:240px;flex:1}
  button.lnk{background:none;border:1px solid var(--line);border-radius:8px;padding:7px 11px;font-size:12px;cursor:pointer;color:var(--fg)}
  button.lnk:hover{background:#eef2f6}
  #cont{color:var(--muted);font-size:12px}
  details{background:var(--card);border:1px solid var(--line);border-radius:10px;margin-bottom:10px;overflow:hidden}
  summary{padding:11px 14px;cursor:pointer;font-weight:600;list-style:none;display:flex;justify-content:space-between;align-items:center;gap:10px}
  summary::-webkit-details-marker{display:none}
  summary:hover{background:#fafbfc}
  summary .caret{color:var(--muted);font-size:12px;transition:transform .15s}
  details[open] summary .caret{transform:rotate(90deg)}
  summary .cnt{font-weight:500;color:var(--muted);font-size:12px}
  table{width:100%;border-collapse:collapse}
  th,td{padding:8px 12px;text-align:left;border-top:1px solid var(--line);font-size:13px;vertical-align:top}
  th{background:#fbfcfd;font-weight:600;color:var(--muted);white-space:nowrap}
  tr.clic:hover{background:#f0f9ff;cursor:pointer}
  .badge{font-size:11px;font-weight:700;padding:2px 7px;border-radius:6px}
  .badge.OK{background:#dcfce7;color:var(--ok)}
  .badge.NO_ENCONTRADO,.badge.ERROR,.badge.DATO_INVALIDO{background:#fef3c7;color:var(--no)}
  .coord{font-variant-numeric:tabular-nums;color:var(--muted);white-space:nowrap}
  a.gm{color:var(--accent);text-decoration:none;white-space:nowrap}
  a.gm:hover{text-decoration:underline}
  .rol{color:var(--muted);font-variant-numeric:tabular-nums}
  footer{color:var(--muted);font-size:12px;text-align:center;padding-bottom:24px}
</style>
</head>
<body>
<header>
  <h1>Consulta de roles &mdash; SII</h1>
  <p>Informaci&oacute;n p&uacute;blica referencial del SII. Generado el __FECHA__.</p>
  <div class="stats">
    <span class="chip">__N_TOTAL__ predios</span>
    <span class="chip ok">__N_OK__ con coordenada</span>
    <span class="chip no">__N_NO__ sin coordenada</span>
    <span class="chip">__N_COM__ comunas</span>
  </div>
</header>
<div id="map"></div>
<div class="wrap">
  <div class="toolbar">
    <input id="q" type="search" placeholder="Filtrar por comuna, direcci&oacute;n o rol...">
    <button class="lnk" id="exp">Expandir todo</button>
    <button class="lnk" id="col">Colapsar todo</button>
    <span id="cont"></span>
  </div>
  <div id="grupos"></div>
</div>
<footer>Clic en una fila para centrar el mapa; &ldquo;Google Maps&rdquo; abre el punto en una pesta&ntilde;a nueva.</footer>

<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"></script>
<script>
const DATOS = __DATOS_JSON__;
const fmt = n => (typeof n === "number") ? n.toFixed(6) : "";
const conCoord = d => d.lat !== null && d.lon !== null;
const gmapUrl = d => "https://www.google.com/maps?q=" + d.lat + "," + d.lon;
const esc = s => (s||"").replace(/[&<>]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));

let map, layer, markers = {};
function initMap(){
  if (typeof L === "undefined"){
    document.getElementById("map").innerHTML =
      '<div style="padding:20px;color:#5b6b7a">No se pudo cargar el mapa (revisa la conexi&oacute;n). La tabla funciona igual.</div>';
    return;
  }
  map = L.map("map", {scrollWheelZoom:true});
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",
    {maxZoom:19, attribution:"&copy; OpenStreetMap"}).addTo(map);
  layer = L.layerGroup().addTo(map);
  const bounds = [];
  DATOS.filter(conCoord).forEach(d => {
    const m = L.marker([d.lat, d.lon]).addTo(layer);
    m.bindPopup(
      "<b>"+esc(d.comuna)+"</b> &middot; rol "+esc(d.manzana)+"-"+esc(d.predio)+"<br>"+
      esc(d.direccion)+"<br><span style='color:#5b6b7a'>"+fmt(d.lat)+", "+fmt(d.lon)+"</span>"+
      "<br><a href='"+gmapUrl(d)+"' target='_blank' rel='noopener'>Abrir en Google Maps</a>");
    markers[d.id] = m; bounds.push([d.lat, d.lon]);
  });
  if (bounds.length) map.fitBounds(bounds, {padding:[30,30], maxZoom:16});
  else map.setView([-34.4,-71.0], 9);
}
function irAPunto(id){
  const m = markers[id];
  if (m && map){ map.flyTo(m.getLatLng(), 18, {duration:.6}); m.openPopup();
    window.scrollTo({top:0, behavior:"smooth"}); }
}
function grupos(filtro){
  const q = (filtro||"").toLowerCase();
  const cont = document.getElementById("grupos");
  cont.innerHTML = "";
  const porComuna = {};
  let vis = 0;
  DATOS.forEach(d => {
    const txt = (d.comuna+" "+d.direccion+" "+d.manzana+"-"+d.predio).toLowerCase();
    if (q && !txt.includes(q)) return;
    (porComuna[d.comuna] = porComuna[d.comuna] || []).push(d); vis++;
  });
  Object.keys(porComuna).sort().forEach(com => {
    const filas = porComuna[com];
    const det = document.createElement("details");
    det.open = !!q;
    let cuerpo = "";
    filas.forEach(d => {
      const cc = conCoord(d);
      const gm = cc ? "<a class='gm' href='"+gmapUrl(d)+"' target='_blank' rel='noopener'>Google Maps</a>" : "";
      cuerpo += "<tr class='"+(cc?"clic":"")+"' data-id='"+d.id+"'>"+
        "<td><span class='badge "+d.estado+"'>"+d.estado.replace(/_/g," ")+"</span></td>"+
        "<td class='rol'>"+esc(d.manzana)+"-"+esc(d.predio)+"</td>"+
        "<td>"+esc(d.direccion)+(d.detalle && d.estado!=="OK" ? "<div style='color:#b45309;font-size:11px'>"+esc(d.detalle)+"</div>" : "")+"</td>"+
        "<td class='coord'>"+(cc? fmt(d.lat)+", "+fmt(d.lon):"")+"</td>"+
        "<td>"+gm+"</td></tr>";
    });
    det.innerHTML =
      "<summary><span><span class='caret'>&#9654;</span> "+esc(com)+"</span>"+
      "<span class='cnt'>"+filas.length+" predio(s)</span></summary>"+
      "<table><thead><tr><th>Estado</th><th>Rol</th><th>Direcci&oacute;n</th>"+
      "<th>Coordenada (lat, lon)</th><th>Mapa</th></tr></thead><tbody>"+cuerpo+"</tbody></table>";
    cont.appendChild(det);
  });
  cont.querySelectorAll("tr.clic").forEach(tr =>
    tr.addEventListener("click", () => irAPunto(+tr.dataset.id)));
  document.getElementById("cont").textContent = vis + " de " + DATOS.length + " predio(s)";
}
document.getElementById("q").addEventListener("input", e => grupos(e.target.value));
document.getElementById("exp").addEventListener("click", () =>
  document.querySelectorAll("#grupos details").forEach(d => d.open = true));
document.getElementById("col").addEventListener("click", () =>
  document.querySelectorAll("#grupos details").forEach(d => d.open = false));
initMap(); grupos("");
</script>
</body>
</html>
"""


def escribir_html(ruta, filas, fecha):
    recs = []
    for i, f in enumerate(filas):
        lat = f.get("lat"); lon = f.get("lon")
        recs.append({
            "id": i,
            "comuna": f.get("comuna") or f.get("consulta_comuna") or "",
            "manzana": f.get("consulta_manzana", ""),
            "predio": f.get("consulta_predio", ""),
            "direccion": f.get("direccion", ""),
            "estado": f.get("estado", "OK"),
            "detalle": f.get("detalle", ""),
            "lat": lat if isinstance(lat, (int, float)) else None,
            "lon": lon if isinstance(lon, (int, float)) else None,
        })
    n_ok = sum(1 for r in recs if r["lat"] is not None and r["lon"] is not None)
    n_no = len(recs) - n_ok
    n_com = len({r["comuna"] for r in recs if r["comuna"]})
    html = (HTML_TEMPLATE
            .replace("__DATOS_JSON__", json.dumps(recs, ensure_ascii=False))
            .replace("__FECHA__", fecha)
            .replace("__N_TOTAL__", str(len(recs)))
            .replace("__N_OK__", str(n_ok))
            .replace("__N_NO__", str(n_no))
            .replace("__N_COM__", str(n_com)))
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(html)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def cmd_list_comunas(filtro):
    n = _norm(filtro) if filtro else ""
    hallados = 0
    for c in sorted(_COMUNAS, key=lambda c: c["nombre"]):
        if not n or n in _norm(c["nombre"]):
            print("  %-7s  %s" % (c["codigo"], c["nombre"]))
            hallados += 1
    print("\n%d comuna(s) coinciden con %r." % (hallados, filtro) if filtro
          else "\nTotal: %d comunas." % hallados)


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Consulta roles de avaluo en el mapa del SII (comuna + manzana + predio).")
    p.add_argument("entrada", nargs="?", help="CSV de entrada con columnas COMUNA, MANZANA, PREDIO.")
    p.add_argument("--comuna", help="Nombre o codigo SII de la comuna (modo consulta unica).")
    p.add_argument("--manzana", help="Numero de manzana (modo consulta unica).")
    p.add_argument("--predio", help="Numero de predio (modo consulta unica).")
    p.add_argument("--rol", help="Alternativa a --manzana/--predio: rol como manzana-predio, ej. 1-1.")
    p.add_argument("-o", "--nombre", default="salida_sii",
                   help="Nombre base de las salidas (sin extension). Por defecto 'salida_sii'.")
    p.add_argument("--outdir", default=DIR_SALIDAS,
                   help="Carpeta donde guardar las salidas (por defecto outputs).")
    p.add_argument("--no-csv", action="store_true", help="No generar el CSV.")
    p.add_argument("--no-xlsx", action="store_true", help="No generar el Excel (.xlsx).")
    p.add_argument("--no-geojson", action="store_true", help="No generar el GeoJSON.")
    p.add_argument("--no-html", action="store_true", help="No generar el HTML.")
    p.add_argument("--delay", type=float, default=0.5, help="Pausa (s) entre consultas.")
    p.add_argument("--reintentos", type=int, default=3, help="Reintentos ante error de red.")
    p.add_argument("--col-comuna", help="Forzar nombre de la columna de comuna.")
    p.add_argument("--col-manzana", help="Forzar nombre de la columna de manzana.")
    p.add_argument("--col-predio", help="Forzar nombre de la columna de predio.")
    p.add_argument("--list-comunas", nargs="?", const="", default=None,
                   help="Lista los codigos de comuna (opcional: filtro) y termina.")
    args = p.parse_args(argv)

    if args.list_comunas is not None:
        cmd_list_comunas(args.list_comunas)
        return 0

    # --- lista de trabajo (comuna, manzana, predio) ------------------------
    trabajo = []
    if args.comuna and (args.rol or (args.manzana and args.predio)):
        if args.rol:
            partes = [x for x in str(args.rol).replace("/", "-").replace(" ", "-").split("-") if x]
            if len(partes) != 2:
                p.error("--rol debe ser manzana-predio, ej. 1-1")
            trabajo.append((args.comuna, partes[0], partes[1]))
        else:
            trabajo.append((args.comuna, args.manzana, args.predio))
    elif args.entrada:
        try:
            trabajo = leer_entrada(args.entrada, args.col_comuna, args.col_manzana, args.col_predio)
        except FileNotFoundError:
            p.error("No se encontro el archivo de entrada: %s" % args.entrada)
        if not trabajo:
            p.error("El archivo no tiene filas utiles (se esperan columnas COMUNA, MANZANA, PREDIO).")
    else:
        p.error("Indica --comuna con --manzana/--predio (o --rol), o un CSV de entrada. "
                "Usa --list-comunas para ver codigos.")

    print("Consultando %d predio(s) en el mapa del SII...\n" % len(trabajo), file=sys.stderr)

    session = requests.Session()
    resultados = []
    ok = sin = err = 0

    for i, (comuna_in, manzana_in, predio_in) in enumerate(trabajo, 1):
        try:
            comuna_cod, comuna_nom = resolver_comuna(comuna_in)
            manzana = _entero(manzana_in, "Manzana")
            predio = _entero(predio_in, "Predio")
        except ValueError as e:
            resultados.append(fila_resultado(comuna_in, manzana_in, predio_in, "", None, "DATO_INVALIDO", str(e)))
            err += 1
            print("  [%d/%d] %-22s %s-%s  -> DATO INVALIDO: %s" %
                  (i, len(trabajo), comuna_in, manzana_in, predio_in, e), file=sys.stderr)
            continue

        try:
            data = consultar(session, comuna_cod, manzana, predio, args.reintentos)
        except RuntimeError as e:
            resultados.append(fila_resultado(comuna_in, manzana_in, predio_in, comuna_nom, None, "ERROR", str(e)))
            err += 1
            print("  [%d/%d] %-22s %d-%d  -> ERROR: %s" %
                  (i, len(trabajo), comuna_nom, manzana, predio, e), file=sys.stderr)
            continue

        if data:
            resultados.append(fila_resultado(comuna_in, manzana_in, predio_in, comuna_nom, data, "OK"))
            ok += 1
            print("  [%d/%d] %-22s %d-%d  -> %s" %
                  (i, len(trabajo), comuna_nom, manzana, predio, (data.get("direccion") or "").strip()),
                  file=sys.stderr)
        else:
            resultados.append(fila_resultado(comuna_in, manzana_in, predio_in, comuna_nom, None,
                                             "NO_ENCONTRADO", "El SII no tiene datos para ese predio."))
            sin += 1
            print("  [%d/%d] %-22s %d-%d  -> sin datos" %
                  (i, len(trabajo), comuna_nom, manzana, predio), file=sys.stderr)

        if i < len(trabajo) and args.delay > 0:
            time.sleep(args.delay)

    # --- escribir salidas en la carpeta outputs ----------------------------
    os.makedirs(args.outdir, exist_ok=True)
    base = os.path.join(args.outdir, os.path.splitext(os.path.basename(args.nombre))[0])
    fecha = time.strftime("%d-%m-%Y %H:%M")

    print("\nListo. %d con coordenada, %d sin datos, %d con error." % (ok, sin, err), file=sys.stderr)
    print("Carpeta de salidas: %s" % os.path.abspath(args.outdir), file=sys.stderr)

    if not args.no_csv:
        escribir_csv(base + ".csv", resultados)
        print("  - CSV:     %s.csv  (para QGIS/scripts; punto decimal)" % os.path.basename(base), file=sys.stderr)
    if not args.no_xlsx:
        if escribir_xlsx(base + ".xlsx", resultados):
            print("  - Excel:   %s.xlsx (coordenadas a prueba de Excel)" % os.path.basename(base), file=sys.stderr)
        else:
            print("  - Excel:   omitido (instala openpyxl:  pip install openpyxl)", file=sys.stderr)
    if not args.no_geojson:
        n = escribir_geojson(base + ".geojson", resultados)
        print("  - GeoJSON: %s.geojson (%d puntos)" % (os.path.basename(base), n), file=sys.stderr)
    if not args.no_html:
        escribir_html(base + ".html", resultados, fecha)
        print("  - HTML:    %s.html (mapa + tabla abatible)" % os.path.basename(base), file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
