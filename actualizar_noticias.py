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

import calendar
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
MAX_FECHAS_A_VERIFICAR = 300  # notas que abre para leer su fecha real (las que no se pueden fechar no se publican)

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

# Palabras extra para reconocer notas sobre CABA en los medios NACIONALES
# (se suman a las de jurisdicciones.json, sin reemplazarlas).
EXTRA_CABA = {
    "nombres": ["Gobierno de la Ciudad", "Legislatura de la Ciudad", "Policía de la Ciudad",
                "Banco Ciudad", "Palermo", "Recoleta", "Puerto Madero", "Barrio Mugica"],
    "palabras": ["jefe de Gobierno", "jefa de Gobierno", "subte", "premetro", "gcba"],
    "personas": ["Gabriel Sánchez Zinny", "Sánchez Zinny", "Gabino Tapia", "Leandro Santoro"],
}
# "la Ciudad" con mayúscula, sin "de"/"del" después: en los medios nacionales
# casi siempre es Buenos Aires ("la Ciudad habilitó...", "en la Ciudad").
# Se descarta "la Ciudad de Córdoba", "Ciudad del Este", etc.
LA_CIUDAD = re.compile(r"(?<![A-Za-z])[Ll]a Ciudad(?![A-Za-z])(?!\s+(?:de|del)\b)")
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
    # Fecha con hora y zona (2026-11-01T23:30:00Z): pasarla a hora argentina,
    # así una nota de anoche a las 22 h no aparece como de hoy.
    m = re.search(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})", valor)
    if m:
        try:
            iso = m.group(0).replace("Z", "+00:00")
            if re.search(r"[+-]\d{4}$", iso):
                iso = iso[:-2] + ":" + iso[-2:]
            d = datetime.fromisoformat(iso).astimezone(ZONA).date()
            return fecha_valida(d.year, d.month, d.day)
        except ValueError:
            pass
    m = re.search(r"(20\d{2})-(\d{2})-(\d{2})", valor)
    if m:
        return fecha_valida(*m.groups())
    m = re.search(r"(\d{1,2})/(\d{1,2})/(20\d{2})", valor)
    if m:
        return fecha_valida(m.group(3), m.group(2), m.group(1))
    return None


_MESES_RE = "|".join(["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                       "agosto", "septiembre", "setiembre", "octubre", "noviembre", "diciembre"])
_FECHA_LARGA = re.compile(r"\b(\d{1,2})\s+de\s+(" + _MESES_RE + r")(?:\s+de|,)?\s+(20\d{2})", re.I)
_FECHA_CORTA = re.compile(r"\b(\d{1,2})[/.-](\d{1,2})[/.-](20\d{2})\b")


def fechas_escritas(texto):
    """Fechas escritas a mano en el texto: '25 de octubre de 2026', '02/11/2026'.
    Devuelve [(posición, fecha)]."""
    salida = []
    for m in _FECHA_LARGA.finditer(texto):
        mes = m.group(2).lower().replace("setiembre", "septiembre")
        f = fecha_valida(m.group(3), MESES.index(mes) + 1, m.group(1))
        if f:
            salida.append((m.start(), f))
    for m in _FECHA_CORTA.finditer(texto):
        f = fecha_valida(m.group(3), m.group(2), m.group(1))
        if f:
            salida.append((m.start(), f))
    return salida


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
                # feedparser da la hora en UTC: pasarla a hora argentina
                d = datetime.fromtimestamp(calendar.timegm(t), ZONA).date()
                f = fecha_valida(d.year, d.month, d.day)
                break
        de_url = fecha_desde_url(link)
        notas.append({"titulo": titulo, "url": urljoin(medio_url, link),
                      "fecha": de_url or f, "origen": "url" if de_url else "rss"})
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
        origen = "url" if f else "portada"
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
                # Fecha escrita junto al titular ("Fecha de publicación: 01/11/2026").
                # Solo si hay una sola fecha distinta: si hay varias, el bloque
                # ya abarca otras notas y no se sabe cuál es cuál.
                distintas = {fe for _, fe in fechas_escritas(padre.get_text(" "))}
                if len(distintas) == 1:
                    f = distintas.pop()
                    break
                if len(distintas) > 1:
                    break
        vistas.add(url)
        notas.append({"titulo": titulo, "url": url, "fecha": f, "origen": origen})
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


def fecha_real_de_nota(url, titulo=""):
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
        r'<time[^>]*itemprop=["\']datePublished["\'][^>]*datetime=["\']([^"\']+)',
        r'<time[^>]*pubdate[^>]*datetime=["\']([^"\']+)',
    ]
    # (No se usa cualquier <time> de la página: muchos medios ponen la fecha
    # de HOY en el encabezado, y eso hacía pasar notas viejas como nuevas.)
    for p in patrones:
        m = re.search(p, html)
        if m:
            f = fecha_desde_texto(m.group(1))
            if f:
                return f
    # Último recurso: la fecha escrita pegada al título de la nota (<h1>),
    # como hace la página del Gobierno porteño ("Domingo 25 de octubre de 2026").
    # Se toma la más cercana al título, para no confundirla con la fecha de hoy
    # del encabezado del sitio ni con las de "últimas noticias" al pie.
    # El <h1> de la nota: si hay varios (logo, nombre del sitio), el que más se
    # parece al título que ya conocemos.
    h1 = None
    clave_titulo = palabras_clave(titulo)
    mejor = 0.0
    for m in re.finditer(r"<h1[\s>](.*?)</h1>", html, re.I | re.S):
        texto = re.sub(r"<[^>]+>", " ", m.group(1))
        parecido = len(clave_titulo & palabras_clave(texto)) / max(1, len(clave_titulo))
        if h1 is None or parecido > mejor:
            h1, mejor = m, parecido
    if h1:
        inicio = max(0, h1.start() - 1500)
        trozo = re.sub(r"<[^>]+>", lambda t: " " * len(t.group(0)), html[inicio:h1.start() + 800])
        antes, despues = [], []
        for pos, f in fechas_escritas(trozo):
            distancia = pos + inicio - h1.start()
            (antes if distancia < 0 else despues).append((abs(distancia), f))
        # Primero la fecha justo arriba del título (volanta/fecha); si no hay,
        # la que está justo debajo (línea de autor). No se lee más lejos para no
        # tomar fechas mencionadas en el cuerpo ("a partir del 2 de noviembre").
        for lista in (antes, despues):
            if lista:
                return min(lista)[1]
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
        if clave == "CABA" and LA_CIUDAD.search(sin_tildes(titulo)):
            hits += 2
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
    if "CABA" in JUR:
        for campo, lista in EXTRA_CABA.items():
            actuales = JUR["CABA"].setdefault(campo, [])
            actuales += [x for x in lista if x not in actuales]
    fuentes = cargar_fuentes()
    historial = {}
    fuentes_vistas = {}
    if ARCHIVO_HISTORIAL.exists():
        historial = json.loads(ARCHIVO_HISTORIAL.read_text(encoding="utf-8"))
    # Medios que ya se venían leyendo en corridas anteriores (y desde cuándo)
    fuentes_vistas = historial.pop("__fuentes__", {})
    motivos = defaultdict(lambda: defaultdict(int))  # por qué se descartó cada nota

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
        # ¿Es la primera vez que leemos este medio? Si ya lo leíamos antes, una
        # nota que aparece por primera vez es nueva de verdad (salió hoy).
        medio_conocido = fuente["url"] in fuentes_vistas
        fuentes_vistas.setdefault(fuente["url"], HOY.isoformat())
        if fuente["bloque"] in ("2", "3") and fuente["juris"] not in ("NACIONAL", "Varias"):
            medio = f"{medio} de {fuente['juris']}"

        for n in notas:
            # Fecha: la de la nota si se conoce; si no, el día en que la vimos por primera vez
            if n["url"] not in historial:
                historial[n["url"]] = {"visto": HOY.isoformat(), "fecha": None,
                                       "nueva": medio_conocido}
            h = historial[n["url"]]
            if n["fecha"] and h.get("origen") != "nota":
                h["fecha"] = n["fecha"].isoformat()
                h["origen"] = n.get("origen")
            fecha = date.fromisoformat(h["fecha"] or h["visto"])
            if (HOY - fecha).days > DIAS_VENTANA:
                motivos[fuente["juris"]]["nota vieja"] += 1
                continue

            puntos = puntaje_politico(n["titulo"], n["url"])
            if puntos <= -10:
                motivos[fuente["juris"]]["tema excluido"] += 1
                continue  # deportes, farándula, espectáculos, policiales, etc.
            if fuente["bloque"] in ("2", "3"):
                puntos += 2      # prensa oficial y legislaturas: hechos de gestión
            elif fuente["bloque"] == "1":
                puntos += 1

            # ¿A qué jurisdicción va?
            if fuente["bloque"] == "6":
                # Columnas de opinión: se eligieron a mano y van solo a Nación
                destinos = ["NACIONAL"]
                puntos += 2
            elif fuente["juris"] in JUR and fuente["juris"] != "NACIONAL" and fuente["bloque"] != "1":
                destinos = [fuente["juris"]]
                if fuente["bloque"] == "5":
                    menciones = jurisdicciones_mencionadas(n["titulo"], JUR)
                    if fuente["juris"] in menciones:
                        puntos += 1
            else:  # medios nacionales, Congreso, ministerios, agregadores
                menciones = jurisdicciones_mencionadas(n["titulo"], JUR)
                if fuente["juris"] == "Menciones" and not menciones:
                    continue  # "lo último" de un diario nacional: solo lo que nombra a una provincia o CABA
                destinos = sorted(menciones, key=menciones.get, reverse=True)[:2] or ["NACIONAL"]
                if destinos != ["NACIONAL"]:
                    puntos += 1

            for destino in destinos:
                criterio, extra = criterio_caba(n["titulo"]) if destino == "CABA" else (None, 0)
                candidatas.append({
                    "juris": destino, "titulo": n["titulo"], "url": n["url"],
                    "medio": medio, "fecha": fecha, "fecha_segura": bool(h["fecha"]) and h.get("origen") in ("url", "nota"),
                    "fecha_listado": bool(h["fecha"]),
                    "nueva": bool(h.get("nueva")) and h["visto"] == HOY.isoformat(),
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
            if len(lista) >= maximo:
                continue
            if e["puntaje"] < (2 if e["imagen"] else PUNTAJE_MINIMO):
                motivos[clave]["puntaje bajo"] += 1
                continue
            if por_medio[rep["medio"]] >= max_mismo_medio:
                continue
            # Si no sabemos la fecha real, abrir la nota para confirmarla.
            # Si no se puede confirmar, la nota NO se publica: antes pasaba con
            # la fecha del día en que la vimos, y así se colaban notas viejas.
            # Excepción: si el medio ya se venía leyendo y la nota apareció
            # recién hoy, es nueva aunque la página no muestre su fecha.
            if not rep["fecha_segura"]:
                f = None
                if verificadas < MAX_FECHAS_A_VERIFICAR:
                    verificadas += 1
                    f = fecha_real_de_nota(rep["url"], rep["titulo"])
                if f:
                    historial[rep["url"]]["fecha"] = f.isoformat()
                    historial[rep["url"]]["origen"] = "nota"
                    rep["fecha"] = f
                    rep["fecha_segura"] = True
                    if (HOY - f).days > DIAS_VENTANA or f > HOY:
                        motivos[clave]["nota vieja (fecha leída en la nota)"] += 1
                        continue
                elif not (rep["fecha_listado"] or rep["nueva"]):
                    # (si la portada o el RSS daban una fecha de hoy y la nota no
                    # se pudo abrir, se usa esa)
                    motivos[clave]["sin fecha confirmable"] += 1
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
    historial["__fuentes__"] = fuentes_vistas
    ARCHIVO_HISTORIAL.write_text(json.dumps(historial, ensure_ascii=False), encoding="utf-8")

    total = sum(len(v["top3"]) + len(v["imagen"]) for v in items.values())
    vacias = [c for c in JUR if c not in items]
    if vacias:
        print("\nJurisdicciones sin noticias y por qué se descartaron sus notas:")
        for c in vacias:
            detalle = ", ".join(f"{m}: {k}" for m, k in motivos[c].items()) or "no llegó ninguna nota"
            print(f"  {c:<20} {detalle}")
    print(f"\nListo: {total} noticias en {len(items)} jurisdicciones. "
          f"Fuentes que no abrieron: {len(caidas)} de {len(fuentes)}.")


if __name__ == "__main__":
    main()
