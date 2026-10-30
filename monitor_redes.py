# -*- coding: utf-8 -*-
"""
MONITOR DE REDES — corre cada hora y guarda "redes.json", que la página
muestra en la ficha de CABA ("Pulso en redes").

Fuentes, todas gratuitas:
  1. Tendencias de X en Argentina (desde trends24.in). No necesita nada.
  2. Búsquedas en alza en Google Argentina (Google Trends). No necesita nada.
  3. Videos de YouTube que mencionan a los temas. Necesita una clave gratuita
     de Google guardada en GitHub como secreto YOUTUBE_API_KEY (opcional).
  4. Posts de Bluesky que mencionan a los temas. Necesita una cuenta gratuita
     de Bluesky: secretos BSKY_HANDLE y BSKY_APP_PASSWORD (opcional).

Si falta una clave, esa fuente se saltea y el resto funciona igual.
X e Instagram no permiten leer posts gratis, por eso no están.
"""

import json
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import feedparser
import requests
from bs4 import BeautifulSoup

from actualizar_noticias import NAVEGADOR, ZONA, contiene, esta_excluida, normalizar

CARPETA = Path(__file__).resolve().parent
ARCHIVO_SALIDA = CARPETA / "redes.json"
ARCHIVO_JURIS = CARPETA / "jurisdicciones.json"
HORAS_GUARDADAS = 48      # cuánto tiempo quedan en la página los videos y posts
MAX_POR_LISTA = 25

# ----------------------------------------------------------------------------
# TEMAS QUE SE MAPEAN (editá acá para sumar personas o temas)
# ----------------------------------------------------------------------------
TEMAS = {
    "Jorge Macri": ["Jorge Macri", "jorgemacri"],
    "Larreta": ["Rodríguez Larreta", "Larreta", "horaciorlarreta"],
    "Ciudad": ["Ciudad de Buenos Aires", "CABA", "porteño", "porteña", "porteños",
               "Legislatura porteña", "Gobierno porteño", "GCBA", "Policía de la Ciudad"],
}
# Para marcar una tendencia como "relevante" (política, gestión, Ciudad).
# Son comienzos de palabra; las de 4 letras o menos tienen que aparecer completas.
PALABRAS_TENDENCIA = [
    "milei", "karina", "kicillof", "cristina", "cfk", "kirchner", "macri", "bullrich",
    "villarruel", "caputo", "adorni", "francos", "santilli", "massa", "lousteau",
    "santoro", "grabois", "espert", "pettovello", "quiros", "libertad avanza", "lla",
    "peronis", "kirchneris", "radical", "ucr", "gobierno", "congreso", "diputad",
    "senad", "legislatura", "ley", "decreto", "dnu", "veto", "presupuesto", "eleccion",
    "electoral", "ministr", "gabinete", "renunci", "paro", "cgt", "ate", "subte",
    "colectivo", "tarifa", "inflacion", "indec", "fmi", "jubilad", "universidad",
    "marcha", "piquete", "inseguridad", "policia", "intendent", "gobernador",
    "oposicion", "oficialismo", "2027",
]

# Lo que se busca en YouTube y Bluesky
BUSQUEDAS = ['"Jorge Macri"', "Larreta", '"Ciudad de Buenos Aires"', "CABA", '"Legislatura porteña"']

AHORA = datetime.now(timezone.utc)


# ----------------------------------------------------------------------------
# ¿DE QUÉ HABLA Y NOS SIRVE?
# ----------------------------------------------------------------------------
def separar_hashtag(texto):
    """#ParoDeSubte -> Paro De Subte"""
    t = texto.lstrip("#")
    return re.sub(r"(?<=[a-záéíóúñ])(?=[A-ZÁÉÍÓÚÑ])", " ", t)


def temas_mencionados(texto):
    t = normalizar(separar_hashtag(texto))
    encontrados = []
    for tema, nombres in TEMAS.items():
        if any(re.search(r"(?<![a-z0-9])" + re.escape(normalizar(n)) + r"(?![a-z0-9])", t) for n in nombres):
            encontrados.append(tema)
    return encontrados


def cargar_palabras_ciudad():
    try:
        J = json.loads(ARCHIVO_JURIS.read_text(encoding="utf-8"))
        caba = J.get("CABA", {})
        return [normalizar(x) for x in caba.get("nombres", []) + caba.get("palabras", []) + caba.get("personas", [])]
    except Exception:
        return []


PALABRAS_CIUDAD = cargar_palabras_ciudad()


def es_relevante(texto):
    """Política, gestión o temas de la Ciudad."""
    if esta_excluida(texto, ""):
        return False
    t = normalizar(separar_hashtag(texto))
    if temas_mencionados(texto):
        return True
    if any(re.search(r"(?<![a-z0-9])" + re.escape(p) + r"(?![a-z0-9])", t) for p in PALABRAS_CIUDAD):
        return True
    return any(contiene(t, w) for w in PALABRAS_TENDENCIA)


# ----------------------------------------------------------------------------
# 1. TENDENCIAS DE X (Argentina)
# ----------------------------------------------------------------------------
def tendencias_x():
    r = requests.get("https://trends24.in/argentina/", headers=NAVEGADOR, timeout=20)
    r.raise_for_status()
    sopa = BeautifulSoup(r.text, "html.parser")
    lista = sopa.select("ol.trend-card__list") or sopa.find_all("ol")
    vistos, salida = set(), []
    # La primera lista es la de la última hora
    for a in (lista[0].find_all("a") if lista else sopa.find_all("a", href=re.compile(r"/search\?q="))):
        tema = a.get_text(" ", strip=True)
        if not tema or tema.lower() in vistos:
            continue
        vistos.add(tema.lower())
        salida.append({
            "tema": tema,
            "url": "https://x.com/search?q=" + quote(tema),
            "relevante": es_relevante(tema),
            "temas": temas_mencionados(tema),
        })
        if len(salida) >= 20:
            break
    if not salida:
        raise ValueError("no se encontraron tendencias en la página")
    return salida


# ----------------------------------------------------------------------------
# 2. BÚSQUEDAS EN ALZA EN GOOGLE (Argentina)
# ----------------------------------------------------------------------------
def tendencias_google():
    feed = feedparser.parse(requests.get("https://trends.google.com/trending/rss?geo=AR",
                                         headers=NAVEGADOR, timeout=20).content)
    salida = []
    for e in feed.entries[:25]:
        titulo = e.get("title", "").strip()
        nota_titulo = e.get("ht_news_item_title", "")
        nota_url = e.get("ht_news_item_url", "")
        texto = f"{titulo} {nota_titulo}"
        salida.append({
            "tema": titulo,
            "busquedas": e.get("ht_approx_traffic", ""),
            "nota": nota_titulo, "nota_url": nota_url,
            "medio": e.get("ht_news_item_source", ""),
            "relevante": es_relevante(texto),
            "temas": temas_mencionados(texto),
        })
    if not salida:
        raise ValueError("Google Trends no devolvió resultados")
    return salida


# ----------------------------------------------------------------------------
# 3. YOUTUBE
# ----------------------------------------------------------------------------
def videos_youtube():
    clave = os.environ.get("YOUTUBE_API_KEY", "").strip()
    if not clave:
        return None
    desde = (AHORA - timedelta(hours=3)).strftime("%Y-%m-%dT%H:%M:%SZ")
    r = requests.get("https://www.googleapis.com/youtube/v3/search", timeout=20, params={
        "key": clave, "part": "snippet", "type": "video", "order": "date", "maxResults": 25,
        "q": "|".join(BUSQUEDAS), "publishedAfter": desde,
        "regionCode": "AR", "relevanceLanguage": "es",
    })
    r.raise_for_status()
    salida = []
    for it in r.json().get("items", []):
        sn = it.get("snippet", {})
        texto = f"{sn.get('title', '')} {sn.get('description', '')}"
        temas = temas_mencionados(texto)
        if not temas or esta_excluida(sn.get("title", ""), ""):
            continue
        salida.append({
            "id": "yt:" + it["id"]["videoId"],
            "titulo": BeautifulSoup(sn.get("title", ""), "html.parser").get_text(),
            "autor": sn.get("channelTitle", ""),
            "fecha": sn.get("publishedAt", ""),
            "url": "https://www.youtube.com/watch?v=" + it["id"]["videoId"],
            "temas": temas,
        })
    return salida


# ----------------------------------------------------------------------------
# 4. BLUESKY
# ----------------------------------------------------------------------------
def posts_bluesky():
    usuario = os.environ.get("BSKY_HANDLE", "").strip()
    clave = os.environ.get("BSKY_APP_PASSWORD", "").strip()
    if not usuario or not clave:
        return None
    sesion = requests.post("https://bsky.social/xrpc/com.atproto.server.createSession", timeout=20,
                           json={"identifier": usuario, "password": clave})
    sesion.raise_for_status()
    token = sesion.json()["accessJwt"]
    desde = (AHORA - timedelta(hours=3)).strftime("%Y-%m-%dT%H:%M:%SZ")
    salida, vistos = [], set()
    for q in BUSQUEDAS:
        r = requests.get("https://bsky.social/xrpc/app.bsky.feed.searchPosts", timeout=20,
                         headers={"Authorization": f"Bearer {token}"},
                         params={"q": q, "sort": "latest", "limit": 25, "since": desde})
        if r.status_code != 200:
            continue
        for p in r.json().get("posts", []):
            if p["uri"] in vistos:
                continue
            vistos.add(p["uri"])
            texto = (p.get("record") or {}).get("text", "")
            temas = temas_mencionados(texto)
            if not temas or esta_excluida(texto, ""):
                continue
            autor = p.get("author", {})
            salida.append({
                "id": "bsky:" + p["uri"],
                "titulo": texto[:280] + ("…" if len(texto) > 280 else ""),
                "autor": "@" + autor.get("handle", ""),
                "fecha": (p.get("record") or {}).get("createdAt", p.get("indexedAt", "")),
                "url": f"https://bsky.app/profile/{autor.get('handle', '')}/post/{p['uri'].rsplit('/', 1)[-1]}",
                "temas": temas,
                "likes": p.get("likeCount", 0),
            })
    return salida


# ----------------------------------------------------------------------------
# PROGRAMA PRINCIPAL
# ----------------------------------------------------------------------------
def acumular(anteriores, nuevos):
    """Junta lo nuevo con lo de las últimas horas, sin repetir, más reciente primero."""
    por_id = {x["id"]: x for x in anteriores}
    for x in nuevos:
        por_id[x["id"]] = x
    limite = (AHORA - timedelta(hours=HORAS_GUARDADAS)).strftime("%Y-%m-%dT%H:%M:%S")
    lista = [x for x in por_id.values() if x.get("fecha", "") >= limite]
    lista.sort(key=lambda x: x.get("fecha", ""), reverse=True)
    return lista[:MAX_POR_LISTA]


def main():
    anterior = {}
    if ARCHIVO_SALIDA.exists():
        try:
            anterior = json.loads(ARCHIVO_SALIDA.read_text(encoding="utf-8"))
        except Exception:
            anterior = {}

    salida = {"actualizado": datetime.now(ZONA).strftime("%d/%m %H:%M"), "fuentes_con_error": []}

    for nombre, funcion in (("tendencias_x", tendencias_x), ("google", tendencias_google)):
        try:
            salida[nombre] = funcion()
            print(f"{nombre}: {len(salida[nombre])} temas ({sum(t['relevante'] for t in salida[nombre])} relevantes)")
        except Exception as e:
            print(f"{nombre}: ERROR {type(e).__name__}: {str(e)[:150]}")
            salida["fuentes_con_error"].append(nombre)
            salida[nombre] = anterior.get(nombre, [])  # si falla, queda lo de la hora anterior

    for nombre, funcion in (("youtube", videos_youtube), ("bluesky", posts_bluesky)):
        try:
            nuevos = funcion()
            if nuevos is None:
                print(f"{nombre}: sin clave configurada, se saltea")
                salida[nombre] = []
                continue
            salida[nombre] = acumular(anterior.get(nombre, []), nuevos)
            print(f"{nombre}: {len(nuevos)} nuevos, {len(salida[nombre])} en las últimas {HORAS_GUARDADAS} h")
        except Exception as e:
            print(f"{nombre}: ERROR {type(e).__name__}: {str(e)[:150]}")
            salida["fuentes_con_error"].append(nombre)
            salida[nombre] = anterior.get(nombre, [])

    ARCHIVO_SALIDA.write_text(json.dumps(salida, ensure_ascii=False, indent=1), encoding="utf-8")
    print("Listo: redes.json actualizado.")


if __name__ == "__main__":
    main()
