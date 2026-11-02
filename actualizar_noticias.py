# -*- coding: utf-8 -*-
"""
ACTUALIZADOR DE NOTICIAS — Mapa político de Argentina (Volvimos)

Qué hace, en pocas palabras:
  1. Lee la lista de medios de "fuentes_limpias.xlsx".
  2. Entra a cada medio y junta los titulares recientes (por RSS cuando el
     medio lo tiene, y si no, leyendo los links de la portada).
  3. Se queda con los que hablan de política o gestión, los asigna a cada
     provincia (según el medio, o según qué provincia/gobernador nombra el título)
     y elige los más relevantes del día (las de días anteriores se descartan).
  4. Guarda el resultado en "noticias.json", que la página lee al abrirse.

No usa inteligencia artificial: elige por palabras clave, cantidad de medios
que cubren el mismo tema y fecha. No escribe "por qué importa".

Para cambiar cosas sin tocar el código, editá:
  - fuentes_limpias.xlsx  -> qué medios se leen
  - jurisdicciones.json   -> nombres, gentilicios y personas de cada provincia
  - las listas de PALABRAS de abajo -> qué cuenta como noticia política
"""

import json
import re
import unicodedata
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, date
from pathlib import Path
from urllib.parse import urljoin, urlparse
from zoneinfo import ZoneInfo

import feedparser
import openpyxl
import requests
from bs4 import BeautifulSoup

# ----------------------------------------------------------------------------
# CONFIGURACIÓN
# ----------------------------------------------------------------------------
CARPETA = Path(__file__).resolve().parent
ARCHIVO_FUENTES = CARPETA / "fuentes_limpias.xlsx"
ARCHIVO_JURIS = CARPETA / "jurisdicciones.json"
ARCHIVO_SALIDA = CARPETA / "noticias.json"
ARCHIVO_HISTORIAL = CARPETA / "historial.json"   # recuerda cuándo vio cada nota

ZONA = ZoneInfo("America/Argentina/Buenos_Aires")
DIAS_VENTANA = 0          # 0 = solo noticias de hoy: al cambiar el día, las de ayer se borran
MAX_TOP = 3               # noticias por jurisdicción (CABA usa MAX_TOP_CABA)
MAX_IMAGEN = 4            # notas de "Imagen, evaluación y candidatos"
PUNTAJE_MINIMO = 3        # debajo de esto, la nota no se publica
MAX_MISMO_MEDIO = 2       # máximo de notas del mismo medio en un top
TIEMPO_ESPERA = 20        # segundos antes de dar por caído un medio
MAX_FECHAS_A_VERIFICAR = 80  # notas que abre para leer su fecha real

NAVEGADOR = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept-Language": "es-AR,es;q=0.9",
}

# Palabras que suman puntos (se comparan sin tildes y en minúscula).
# Cada una es el comienzo de una palabra: "eleccion" cubre "elecciones".
# Las palabras de 4 letras o menos, y las que terminan en "$", tienen que
# aparecer completas ("pro" no cuenta dentro de "proyecto").
PALABRAS_FUERTES = [
    "gobernador", "gobernadora", "vicegobernador", "eleccion", "electoral", "2027",
    "candidat", "gabinete", "renunci", "designa", "design", "ministr", "legislatura",
    "reeleccion", "alianza", "presupuesto", "jura", "asumio", "asumira", "desdobla",
    "boleta", "constituyente", "enmienda", "veto", "decreto",
]
PALABRAS_POLITICAS = [
    "intendent", "interna$", "internas$", "frente", "bloque", "diputad", "senad", "concejal",
    "concejo deliberante", "ley", "proyecto", "sesion", "aprob", "reforma",
    "peronis", "pj", "ucr", "radical", "pro", "libertari", "lla", "milei",
    "casa rosada", "coparticipacion", "endeudamiento", "deuda", "paritaria",
    "gremi", "ate", "docente", "estatales", "oficialismo", "oposicion",
    "provincias unidas", "fuerza patria", "paso", "cronograma", "convencion",
    "gobierno", "funcionari", "justicialis", "kirchner", "cristina", "macri",
    "bullrich", "caputo", "congreso", "fondos", "obra publica", "licitacion",
    "salario", "transferencias", "partido", "lista$", "listas$", "armado",
]
# Notas que van a "Imagen, evaluación y candidatos"
PALABRAS_IMAGEN = [
    "encuesta", "sondeo", "imagen", "intencion de voto", "medicion", "consultora",
]
PALABRAS_CANDIDATURA = ["candidat", "postula", "reeleccion", "encabez"]
PALABRAS_NEGATIVAS = [
    "choque", "accidente", "murio", "muerte de", "asesin", "robo$", "incendio",
    "horoscopo", "receta", "pronostico", "detuvieron",
]
# Secciones que no se leen nunca
SECCIONES_DESCARTADAS = [
    "/policial", "/horoscopo", "/clima", "/tecnologia", "/gastronomia",
    "/turismo", "/autos", "/mundo", "/internacional", "/salud-y-bienestar",
    "/estilo", "/tendencias", "/videos", "/servicios", "/agro",
]

# FILTRO DE EXCLUSIÓN ESTRICTO: deportes, farándula y espectáculos.
# Estas notas se descartan por completo, salvo que el título muestre un
# impacto institucional, judicial o político directo (PALABRAS_INSTITUCIONALES).
SECCIONES_EXCLUIDAS = ["/deporte", "/futbol", "/espectaculo", "/show", "/famosos", "/farandula"]
PALABRAS_EXCLUIDAS = [
    "futbol", "gol$", "goles$", "hincha", "torneo", "superclasico", "river plate",
    "boca juniors", "seleccion argentina", "messi", "tenis$", "formula 1", "farandula",
    "famos", "actriz", "actor$", "cantante", "influencer", "streamer", "gran hermano",
    "recital", "concierto", "estreno", "pelicula", "serie$", "novela", "espectaculo",
]
PALABRAS_INSTITUCIONALES = [
    "legislatura", "gobierno", "justicia", "juez", "jueza", "fiscal", "denuncia",
    "imputad", "procesad", "ley$", "decreto", "clausura", "ministr", "jefe de gobierno",
    "intendent", "gobernador", "congreso", "diputad", "senad", "funcionari",
    "licitacion", "contrato", "subsidio", "presupuesto", "milei", "macri", "kicillof",
]

# ----------------------------------------------------------------------------
# CABA: COBERTURA AMPLIADA
# ----------------------------------------------------------------------------
# En CABA se muestran más noticias y se buscan cinco criterios. Primero entra
# la mejor noticia de cada criterio (si la hay) y después se completa por
# relevancia. Cada noticia lleva su criterio como etiqueta.
MAX_TOP_CABA = 10         # CABA: 10 noticias del día
MAX_MISMO_MEDIO_CABA = 5  # en CABA, hasta 5 notas del mismo medio
CRITERIOS_CABA = {
    # C. Cambios de gabinete (se prioriza prensa oficial y Boletín Oficial)
    "Gabinete": [
        "renunci", "design", "desplaz", "nombr", "asumi", "jura$", "juro$", "gabinete",
        "ministr", "secretari", "jefe de gabinete", "vicejef", "reestructur", "fusion",
        "organigrama", "boletin oficial",
    ],
    # D. Calendario electoral
    "Calendario electoral": [
        "eleccion", "electoral", "desdobla", "concurren", "paso$", "boleta",
        "reforma electoral", "reforma constitucional", "reeleccion", "convocatoria",
        "cronograma", "ficha limpia", "voto$", "votacion",
    ],
    # A. Política
    "Política": [
        "interna$", "internas$", "alianza", "candidat", "armado", "2027", "frente",
        "bloque", "ruptura", "rompi", "quiebre", "tension", "conflicto", "pelea",
        "enfrent", "casa rosada", "milei", "gobierno nacional", "gobernador",
        "libertari", "lla$", "pro$", "ucr$", "peronis", "pj$", "oposicion",
        "oficialismo", "acuerdo",
    ],
    # E. Gestión municipal y problemáticas de la Ciudad
    "Ciudad": [
        "subte", "colectivo", "transporte", "recoleccion", "residuos", "basura",
        "contenedor", "transito", "semaforo", "bache", "bacheo", "asfalto",
        "iluminacion", "luminaria", "vereda", "plaza", "inundac", "anegam", "desague",
        "arbolado", "limpieza", "cuidacoches", "trapitos", "inseguridad", "motochorro",
        "delito", "vecinos", "vecinal", "reclamo", "protesta", "paro$", "gremio",
        "sutecba", "municipales", "metrodelegados", "emova", "clausura", "agc$",
        "situacion de calle", "autopista", "ciclovia", "estacionamiento", "acarreo",
        "grua", "obras$", "obra$",
    ],
    # B. Gestión y políticas públicas
    "Gestión": [
        "legislatura", "ley$", "leyes$", "proyecto", "sesion", "aprob", "decreto",
        "presupuesto", "endeudamiento", "deuda", "impuesto", "abl$", "tasa$", "salud",
        "hospital", "educacion", "escuela", "docente", "seguridad", "policia",
        "obra publica", "economia", "empleo", "alquiler", "vivienda", "licitacion",
    ],
}

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


# ----------------------------------------------------------------------------
# HERRAMIENTAS DE TEXTO
# ----------------------------------------------------------------------------
def sin_tildes(texto):
    texto = unicodedata.normalize("NFD", texto or "")
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def normalizar(texto):
    return re.sub(r"\s+", " ", sin_tildes(texto).lower()).strip()


def contiene(texto_norm, palabra):
    """¿El texto tiene una palabra que empieza con 'palabra'?"""
    completa = palabra.endswith("$") or len(palabra) <= 4
    palabra = palabra.rstrip("$")
    fin = r"(?![a-z0-9])" if completa else ""
    return re.search(r"(?<![a-z0-9])" + re.escape(palabra) + fin, texto_norm) is not None


def contiene_exacto(texto, frase, respetar_mayusculas):
    t = sin_tildes(texto)
    f = sin_tildes(frase)
    flags = 0 if respetar_mayusculas else re.IGNORECASE
    return re.search(r"(?<![A-Za-z0-9])" + re.escape(f) + r"(?![A-Za-z0-9])", t, flags) is not None


def palabras_clave(titulo):
    return {w for w in re.findall(r"[a-z0-9]+", normalizar(titulo)) if len(w) > 3}


# ----------------------------------------------------------------------------
# FECHAS
# ----------------------------------------------------------------------------
HOY = datetime.now(ZONA).date()


def fecha_valida(a, m, d):
    try:
        f = date(int(a), int(m), int(d))
    except ValueError:
        return None
    if HOY - timedelta(days=400) <= f <= HOY + timedelta(days=1):
        return f
    return None


def fecha_desde_url(url):
    """Muchos medios ponen la fecha en el link: /2026/10/21/ o 20261021 o nid21102026."""
    m = re.search(r"/(20\d{2})[/-](\d{1,2})[/-](\d{1,2})(?:/|-|$)", url)
    if m:
        return fecha_valida(*m.groups())
    m = re.search(r"nid(\d{2})(\d{2})(20\d{2})", url)          # La Nación
    if m:
        return fecha_valida(m.group(3), m.group(2), m.group(1))
    m = re.search(r"(?<!\d)(20\d{2})(\d{2})(\d{2})\d{0,8}(?!\d)", url)  # 20261021...
    if m:
        return fecha_valida(*m.groups())
    return None


def fecha_desde_texto(valor):
    if not valor:
        return None
    m = re.search(r"(20\d{2})-(\d{2})-(\d{2})", valor)
    if m:
        return fecha_valida(*m.groups())
    m = re.search(r"(\d{1,2})/(\d{1,2})/(20\d{2})", valor)
    if m:
        return fecha_valida(m.group(3), m.group(2), m.group(1))
    return None


# ----------------------------------------------------------------------------
# LECTURA DE CADA MEDIO
# ----------------------------------------------------------------------------
def descargar(url):
    r = requests.get(url, headers=NAVEGADOR, timeout=TIEMPO_ESPERA)
    r.raise_for_status()
    if "charset" not in r.headers.get("Content-Type", "").lower():
        r.encoding = r.apparent_encoding  # adivina tildes y eñes si el sitio no lo aclara
    return r


def notas_de_rss(contenido, medio_url):
    feed = feedparser.parse(contenido)
    notas = []
    for e in feed.entries[:60]:
        titulo = (e.get("title") or "").strip()
        link = e.get("link") or ""
        if not titulo or not link:
            continue
        f = None
        for campo in ("published_parsed", "updated_parsed"):
            t = e.get(campo)
            if t:
                f = fecha_valida(t.tm_year, t.tm_mon, t.tm_mday)
                break
        notas.append({"titulo": titulo, "url": urljoin(medio_url, link),
                      "fecha": f or fecha_desde_url(link)})
    return notas


def dominio_base(url):
    """infobae.com, clarin.com, cba.gov.ar... (sin www ni subdominios)."""
    partes = urlparse(url).netloc.lower().split(":")[0].split(".")
    n = 3 if len(partes) >= 3 and partes[-2] in ("com", "gob", "gov", "org", "net", "edu") else 2
    return ".".join(partes[-n:])


def notas_de_portada(html, medio_url):
    sopa = BeautifulSoup(html, "html.parser")
    base = dominio_base(medio_url)
    vistas, notas = set(), []
    for a in sopa.find_all("a", href=True):
        url = urljoin(medio_url, a["href"]).split("#")[0]
        if not url.startswith("http") or url in vistas:
            continue
        if dominio_base(url) != base:
            continue  # links a otros sitios
        ruta = urlparse(url).path
        if len(ruta) < 15 or ("-" not in ruta and not re.search(r"\d{4,}", ruta)):
            continue  # links de menú, secciones, etc.
        titulo = a.get_text(" ", strip=True) or a.get("title", "")
        if len(titulo) < 25:
            h = a.find(["h1", "h2", "h3", "h4"])
            titulo = h.get_text(" ", strip=True) if h else titulo
        titulo = re.sub(r"\s+", " ", titulo).strip()
        if len(titulo) < 25 or len(titulo) > 230 or len(titulo.split()) < 5:
            continue
        f = fecha_desde_url(url)
        if not f:  # buscar una etiqueta <time> cerca del link
            padre = a
            for _ in range(3):
                padre = padre.parent
                if padre is None:
                    break
                t = padre.find("time")
                if t is not None:
                    f = fecha_desde_texto(t.get("datetime") or t.get_text())
                    break
        vistas.add(url)
        notas.append({"titulo": titulo, "url": url, "fecha": f})
    return notas[:120]


def leer_medio(fuente):
    """Devuelve (fuente, notas, error)."""
    try:
        r = descargar(fuente["url"])
        tipo = r.headers.get("Content-Type", "")
        if "xml" in tipo or r.text.lstrip().startswith("<?xml"):
            return fuente, notas_de_rss(r.content, fuente["url"]), None
        sopa = BeautifulSoup(r.text, "html.parser")
        # ¿La portada anuncia un RSS? Si lo tiene, es más confiable (trae fechas).
        for link in sopa.find_all("link", attrs={"type": re.compile("rss|atom")})[:2]:
            try:
                rss = descargar(urljoin(fuente["url"], link.get("href", "")))
                notas = notas_de_rss(rss.content, fuente["url"])
                if len(notas) >= 5:
                    return fuente, notas, None
            except Exception:
                pass
        return fuente, notas_de_portada(r.text, fuente["url"]), None
    except Exception as e:
        return fuente, [], f"{type(e).__name__}: {str(e)[:120]}"


def fecha_real_de_nota(url):
    """Abre una nota y busca su fecha de publicación en los datos de la página."""
    try:
        html = descargar(url).text
    except Exception:
        return None
    patrones = [
        r'property=["\']article:published_time["\'][^>]*content=["\']([^"\']+)',
        r'content=["\']([^"\']+)["\'][^>]*property=["\']article:published_time',
        r'"datePublished"\s*:\s*"([^"]+)"',
        r'itemprop=["\']datePublished["\'][^>]*content=["\']([^"\']+)',
        r'<time[^>]*datetime=["\']([^"\']+)',
    ]
    for p in patrones:
        m = re.search(p, html)
        if m:
            f = fecha_desde_texto(m.group(1))
            if f:
                return f
    return None


# ----------------------------------------------------------------------------
# CLASIFICACIÓN Y PUNTAJE
# ----------------------------------------------------------------------------
def esta_excluida(titulo, url):
    """Filtro estricto: deportes, farándula y espectáculos afuera, salvo impacto institucional."""
    t = normalizar(titulo)
    ruta = urlparse(url).path.lower()
    if any(s in ruta for s in SECCIONES_DESCARTADAS):
        return True
    es_de_ocio = any(s in ruta for s in SECCIONES_EXCLUIDAS) or any(contiene(t, w) for w in PALABRAS_EXCLUIDAS)
    if es_de_ocio and not any(contiene(t, w) for w in PALABRAS_INSTITUCIONALES):
        return True
    return False


def criterio_caba(titulo):
    """Devuelve (criterio, puntos extra) para una noticia de CABA."""
    t = normalizar(titulo)
    mejor, mejor_hits, total = "Gestión", 0, 0
    for criterio, palabras in CRITERIOS_CABA.items():
        hits = sum(1 for w in palabras if contiene(t, w))
        total += hits * (2 if criterio == "Ciudad" else 1)
        if hits > mejor_hits:  # en empate gana el que está primero en la lista
            mejor, mejor_hits = criterio, hits
    return mejor, min(total, 5)


def puntaje_politico(titulo, url):
    t = normalizar(titulo)
    ruta = urlparse(url).path.lower()
    if esta_excluida(titulo, url):
        return -10
    p = 0
    p += sum(2 for w in PALABRAS_FUERTES if contiene(t, w))
    p += sum(1 for w in PALABRAS_POLITICAS if contiene(t, w))
    p -= sum(3 for w in PALABRAS_NEGATIVAS if contiene(t, w))
    if "/politica" in ruta:
        p += 2
    return min(p, 10)


def es_de_imagen(titulo):
    t = normalizar(titulo)
    return any(contiene(t, w) for w in PALABRAS_IMAGEN) or (
        contiene(t, "2027") and any(contiene(t, w) for w in PALABRAS_CANDIDATURA))


def jurisdicciones_mencionadas(titulo, JUR):
    encontradas = {}
    for clave, d in JUR.items():
        hits = 0
        hits += sum(2 for n in d.get("nombres", []) if contiene_exacto(titulo, n, True))
        hits += sum(2 for n in d.get("palabras", []) if contiene_exacto(titulo, n, False))
        hits += sum(3 for n in d.get("personas", []) if contiene_exacto(titulo, n, False))
        if hits:
            encontradas[clave] = hits
    return encontradas


def normalizar_jurisdiccion(nombre):
    n = (nombre or "").strip()
    return {"Nación": "NACIONAL", "Tierra del Fuego AIAS": "Tierra del Fuego"}.get(n, n)


# ----------------------------------------------------------------------------
# PROGRAMA PRINCIPAL
# ----------------------------------------------------------------------------
def cargar_fuentes():
    hoja = openpyxl.load_workbook(ARCHIVO_FUENTES, read_only=True).active
    fuentes = []
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if not fila or not fila[3]:
            continue
        bloque, juris, nombre, url = (str(x or "").strip() for x in fila[:4])
        fuentes.append({"bloque": bloque[:1], "juris": normalizar_jurisdiccion(juris),
                        "medio": nombre, "url": url})
    return fuentes


def main():
    JUR = json.loads(ARCHIVO_JURIS.read_text(encoding="utf-8"))
    fuentes = cargar_fuentes()
    historial = {}
    if ARCHIVO_HISTORIAL.exists():
        historial = json.loads(ARCHIVO_HISTORIAL.read_text(encoding="utf-8"))

    print(f"Leyendo {len(fuentes)} fuentes...")
    with ThreadPoolExecutor(max_workers=10) as pool:
        resultados = list(pool.map(leer_medio, fuentes))

    caidas = []
    candidatas = []  # todas las notas útiles, ya asignadas a una jurisdicción
    for fuente, notas, error in resultados:
        estado = f"ERROR ({error})" if error else f"{len(notas)} titulares"
        print(f"  {fuente['juris']:<20} {fuente['medio'][:30]:<30} {estado}")
        if error or not notas:
            caidas.append(fuente["url"])
            continue
        # Nombre del medio: para prensa oficial y legislaturas, aclarar la provincia
        medio = fuente["medio"]
        if fuente["bloque"] in ("2", "3") and fuente["juris"] not in ("NACIONAL", "Varias"):
            medio = f"{medio} de {fuente['juris']}"

        for n in notas:
            # Fecha: la de la nota si se conoce; si no, el día en que la vimos por primera vez
            h = historial.setdefault(n["url"], {"visto": HOY.isoformat(), "fecha": None})
            if n["fecha"]:
                h["fecha"] = n["fecha"].isoformat()
            fecha = date.fromisoformat(h["fecha"] or h["visto"])
            if (HOY - fecha).days > DIAS_VENTANA:
                continue

            puntos = puntaje_politico(n["titulo"], n["url"])
            if puntos <= -10:
                continue  # deportes, farándula, espectáculos, policiales, etc.
            if fuente["bloque"] in ("2", "3"):
                puntos += 2      # prensa oficial y legislaturas: hechos de gestión
            elif fuente["bloque"] == "1":
                puntos += 1

            # ¿A qué jurisdicción va?
            if fuente["juris"] in JUR and fuente["juris"] != "NACIONAL" and fuente["bloque"] != "1":
                destinos = [fuente["juris"]]
                if fuente["bloque"] == "5":
                    menciones = jurisdicciones_mencionadas(n["titulo"], JUR)
                    if fuente["juris"] in menciones:
                        puntos += 1
            else:  # medios nacionales, Congreso, ministerios, agregadores
                menciones = jurisdicciones_mencionadas(n["titulo"], JUR)
                destinos = sorted(menciones, key=menciones.get, reverse=True)[:2] or ["NACIONAL"]
                if destinos != ["NACIONAL"]:
                    puntos += 1

            for destino in destinos:
                criterio, extra = criterio_caba(n["titulo"]) if destino == "CABA" else (None, 0)
                candidatas.append({
                    "juris": destino, "titulo": n["titulo"], "url": n["url"],
                    "medio": medio, "fecha": fecha, "fecha_segura": bool(h["fecha"]),
                    "puntos": puntos + extra, "bloque": fuente["bloque"], "criterio": criterio,
                    "imagen": es_de_imagen(n["titulo"]) or fuente["bloque"] == "4",
                    "local": fuente["juris"] == destino,
                })

    # Agrupar notas que cuentan lo mismo (en la misma jurisdicción)
    por_juris = defaultdict(list)
    for c in candidatas:
        por_juris[c["juris"]].append(c)

    verificadas = 0
    items = {}
    for clave in JUR:
        grupos = []
        for c in sorted(por_juris.get(clave, []), key=lambda x: -x["puntos"]):
            kw = palabras_clave(c["titulo"])
            for g in grupos:
                comunes = len(kw & g["kw"]) / max(1, min(len(kw), len(g["kw"])))
                if c["url"] == g["notas"][0]["url"] or (comunes >= 0.6 and len(kw & g["kw"]) >= 3):
                    g["notas"].append(c)
                    break
            else:
                grupos.append({"kw": kw, "notas": [c]})

        elegidas = []
        for g in grupos:
            medios = []
            for c in g["notas"]:
                if c["medio"] not in medios:
                    medios.append(c["medio"])
            # Representante: preferir fuente oficial, después medio local, después el de más puntos
            rep = sorted(g["notas"], key=lambda c: (c["bloque"] in ("2", "3"), c["local"], c["puntos"]), reverse=True)[0]
            dias = (HOY - rep["fecha"]).days
            puntaje = max(c["puntos"] for c in g["notas"]) + 1.5 * (len(medios) - 1) + max(0, 1.5 - dias * 0.2)
            elegidas.append({"rep": rep, "medios": medios, "puntaje": puntaje,
                             "imagen": any(c["imagen"] for c in g["notas"])})
        elegidas.sort(key=lambda e: -e["puntaje"])
        max_top = MAX_TOP_CABA if clave == "CABA" else MAX_TOP
        if clave == "CABA":
            # Primero la mejor noticia de cada criterio, después el resto por relevancia
            primeras, vistos = [], set()
            for e in elegidas:
                if not e["imagen"] and e["rep"]["criterio"] not in vistos and e["puntaje"] >= PUNTAJE_MINIMO:
                    vistos.add(e["rep"]["criterio"])
                    primeras.append(e)
            elegidas = primeras + [e for e in elegidas if e not in primeras]

        top, imagen, por_medio = [], [], defaultdict(int)
        max_mismo_medio = MAX_MISMO_MEDIO_CABA if clave == "CABA" else MAX_MISMO_MEDIO
        for e in elegidas:
            rep = e["rep"]
            lista, maximo = (imagen, MAX_IMAGEN) if e["imagen"] else (top, max_top)
            if len(lista) >= maximo or e["puntaje"] < (2 if e["imagen"] else PUNTAJE_MINIMO):
                continue
            if por_medio[rep["medio"]] >= max_mismo_medio:
                continue
            # Si no sabemos la fecha real, abrir la nota para confirmarla
            if not rep["fecha_segura"] and verificadas < MAX_FECHAS_A_VERIFICAR:
                verificadas += 1
                f = fecha_real_de_nota(rep["url"])
                if f:
                    historial[rep["url"]]["fecha"] = f.isoformat()
                    rep["fecha"] = f
                    if (HOY - f).days > DIAS_VENTANA:
                        continue
            otros = [m for m in e["medios"] if m != rep["medio"]]
            nota = {"titulo": rep["titulo"], "medio": rep["medio"],
                    "fecha": rep["fecha"].strftime("%d/%m/%Y"), "url": rep["url"]}
            etiquetas = []
            if clave == "CABA" and not e["imagen"]:
                etiquetas.append(rep["criterio"])
            if otros:
                etiquetas.append("También en " + ", ".join(otros[:2]) + (f" y {len(otros) - 2} más" if len(otros) > 2 else ""))
            if etiquetas:
                nota["tag"] = " · ".join(etiquetas)
            lista.append(nota)
            por_medio[rep["medio"]] += 1
        if top or imagen:
            items[clave] = {"top3": top, "imagen": imagen}

    inicio = HOY - timedelta(days=DIAS_VENTANA)
    if DIAS_VENTANA == 0:
        semana = f"{HOY.day} de {MESES[HOY.month - 1]} de {HOY.year}"
    elif inicio.month == HOY.month:
        semana = f"{inicio.day} al {HOY.day} de {MESES[HOY.month - 1]} de {HOY.year}"
    else:
        semana = f"{inicio.day} de {MESES[inicio.month - 1]} al {HOY.day} de {MESES[HOY.month - 1]} de {HOY.year}"

    salida = {
        "semana": semana,
        "fecha": HOY.isoformat(),   # la página oculta las noticias si no son de hoy
        "actualizado": datetime.now(ZONA).strftime("%d/%m %H:%M"),
        "items": items,
        "fuentes_no_accesibles": caidas,
    }
    ARCHIVO_SALIDA.write_text(json.dumps(salida, ensure_ascii=False, indent=1), encoding="utf-8")

    # El historial solo guarda los últimos 30 días, para que no crezca para siempre
    limite = (HOY - timedelta(days=30)).isoformat()
    historial = {u: h for u, h in historial.items() if h["visto"] >= limite}
    ARCHIVO_HISTORIAL.write_text(json.dumps(historial, ensure_ascii=False), encoding="utf-8")

    total = sum(len(v["top3"]) + len(v["imagen"]) for v in items.values())
    print(f"\nListo: {total} noticias en {len(items)} jurisdicciones. "
          f"Fuentes que no abrieron: {len(caidas)} de {len(fuentes)}.")


if __name__ == "__main__":
    main()
