"""
cascada_gaceteros.py — Cascada de georreferenciación multi-gacetero (Manual SiB)
© 2026 Ximena Bedoya Araque · Universidad CES · CBUCES

SISTEMA ABIERTO: acepta una LISTA ORDENADA de gaceteros (rural, urbano, el que sea).
Para cada localidad recorre los gaceteros en orden de especificidad; si ninguno
resuelve con CONFIANZA, baja a municipio → departamento → país → marca (regla de oro).

Fiel al manual (3.5.2): se usa la fuente MÁS ESPECÍFICA que se pueda ubicar, y la
búsqueda del topónimo se ACOTA al municipio/departamento reportado (evita homónimos).

CÓMO AÑADIR UN GACETERO NUEVO (ej. urbano del DANE):
  Solo debe ser un DataFrame con estas columnas (mismo molde que el IGAC):
    nombre, nombre_norm, nombre_compact, departamento, municipio,
    depto_norm, muni_norm, lat, lon   (nombre_alt_norm es opcional)
  Se "enchufa" pasándolo en la lista `gaceteros` de georreferenciar_localidad().
  No hay que tocar el código de la cascada.

Anti-Montelíbano: usa norm_geo (tildes, ñ, espacios) y comparación compacta.
"""

import re
import unicodedata
import difflib
import pandas as pd

UMBRAL_FUZZY = 0.90        # similitud mínima para aceptar coincidencia aproximada
PALABRAS_VACIAS = ("sin datos", "sin informacion", "", "nan", "none")


def norm_geo(t):
    if t is None or (isinstance(t, float) and pd.isna(t)):
        return ""
    t = str(t).strip()
    if t.lower() in ("", "nan", "none"):
        return ""
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn").lower()
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _compacto(s):
    return s.replace(" ", "")


def cargar_gacetero(ruta_parquet):
    """Carga un gacetero en el molde estándar. Verifica columnas mínimas."""
    g = pd.read_parquet(ruta_parquet)
    req = {"nombre_norm", "muni_norm", "lat", "lon"}
    faltan = req - set(g.columns)
    if faltan:
        raise ValueError(f"Al gacetero le faltan columnas: {faltan}")
    if "nombre_compact" not in g.columns:
        g["nombre_compact"] = g["nombre_norm"].map(_compacto)
    return g


def buscar_en_gacetero(localidad, municipio, departamento, gac, nombre_gac=""):
    """
    Busca un topónimo en UN gacetero, acotando al municipio reportado.
    Retorna (lat, lon, detalle) si hay match confiable, o (None, None, motivo).
    Aplica la regla de oro: si es ambiguo (varios puntos), NO elige → None.
    """
    ln = norm_geo(localidad)
    if not ln or ln in PALABRAS_VACIAS or "sin datos" in ln:
        return None, None, "localidad vacía/sin datos"
    lc = _compacto(ln)
    mn = norm_geo(municipio)

    # Acotar al municipio (exacto o compacto, por si el gacetero concatena nombres)
    sub = gac[gac["muni_norm"] == mn]
    if not len(sub):
        sub = gac[gac["muni_norm"].map(_compacto) == _compacto(mn)]
    if not len(sub):
        return None, None, f"municipio '{municipio}' no está en {nombre_gac or 'gacetero'}"

    # 1. Coincidencia exacta del nombre normalizado
    e = sub[sub["nombre_norm"] == ln]
    if len(e) == 1:
        r = e.iloc[0]
        return r["lat"], r["lon"], f"{nombre_gac}: '{r['nombre']}' (exacto)"
    if len(e) > 1:
        prom = _promediar_si_cercanos(e)
        if prom:
            return prom[0], prom[1], f"{nombre_gac}: '{e.iloc[0]['nombre']}' ({len(e)} puntos del mismo sitio)"
        return None, None, f"{len(e)} sitios '{localidad}' en {municipio} (ambiguo, lejanos) → no se elige"

    # 2. Coincidencia compacta (espacios/tildes)
    c = sub[sub["nombre_compact"] == lc]
    if len(c) == 1:
        r = c.iloc[0]
        return r["lat"], r["lon"], f"{nombre_gac}: '{r['nombre']}' (compacto)"
    if len(c) > 1:
        prom = _promediar_si_cercanos(c)
        if prom:
            return prom[0], prom[1], f"{nombre_gac}: '{c.iloc[0]['nombre']}' ({len(c)} puntos del mismo sitio)"
        return None, None, f"{len(c)} sitios compactos (ambiguo, lejanos) → no se elige"

    # 3. La localidad CONTIENE el topónimo (ej. "Hacienda Cuba, margen suroriental")
    #    Solo si el topónimo es suficientemente largo y único en el municipio.
    contiene = sub[sub["nombre_norm"].map(lambda x: len(x) > 5 and x in ln)]
    if len(contiene) == 1:
        r = contiene.iloc[0]
        return r["lat"], r["lon"], f"{nombre_gac}: '{r['nombre']}' (dentro de la descripción)"
    if len(contiene) > 1:
        # quedarnos con el topónimo más largo (más específico) si es único en longitud
        contiene = contiene.copy()
        contiene["_len"] = contiene["nombre_norm"].str.len()
        top = contiene.sort_values("_len", ascending=False)
        if (top["_len"].iloc[0] != top["_len"].iloc[1]):
            r = top.iloc[0]
            return r["lat"], r["lon"], f"{nombre_gac}: '{r['nombre']}' (más específico en la descripción)"
        return None, None, "varios topónimos en la descripción (ambiguo) → no se elige"

    # 4. Aproximada (fuzzy) DENTRO del municipio, umbral alto
    best, br = None, 0.0
    for _, row in sub.iterrows():
        s = difflib.SequenceMatcher(None, ln, row["nombre_norm"]).ratio()
        if s > br:
            br, best = s, row
    if best is not None and br >= UMBRAL_FUZZY:
        return best["lat"], best["lon"], f"{nombre_gac}: '{best['nombre']}' (aproximado {br:.0%})"

    return None, None, f"no se encontró '{localidad}' en {municipio}"


def georreferenciar_localidad(localidad, municipio, departamento, nivel,
                              gaceteros, geo_idx, match_municipio_fn,
                              centroide_depto_fn=None, centroide_pais=(4.5709, -74.2973)):
    """
    CASCADA COMPLETA fiel al manual. Recorre los gaceteros en orden y baja de nivel
    si no hay match confiable.

    gaceteros: lista de tuplas (nombre, DataFrame) en orden de especificidad.
               Ej: [("IGAC rural", df_igac), ("DANE urbano", df_dane)]
    geo_idx:   índice GADM de geo_match (para centroides municipio/depto).
    match_municipio_fn / centroide_depto_fn: funciones de geo_match.

    Retorna dict: {lat, lon, fuente, metodo, estado, detalle}.
    """
    # Nivel 7: NO se georreferencia (manual)
    if nivel == 7:
        return {"lat": None, "lon": None, "fuente": "", "estado": "Nivel 7",
                "metodo": "No se georreferencia (información dudosa/contradictoria)",
                "detalle": "Requiere revisión del curador."}

    # 1-2. CASCADA DE GACETEROS (topónimo específico)
    for nombre_gac, gac in gaceteros:
        lat, lon, det = buscar_en_gacetero(localidad, municipio, departamento, gac, nombre_gac)
        if lat is not None:
            return {"lat": float(lat), "lon": float(lon), "fuente": nombre_gac,
                    "estado": "[✓] OK",
                    "metodo": f"topónimo localizado en {nombre_gac}",
                    "detalle": det}

    # 3. CENTROIDE DE MUNICIPIO (GADM) — para niveles 1-4 (1 entra aquí si su coord falló)
    if nivel in (1, 2, 3, 4) or nivel is None:
        mlat, mlon, minfo = match_municipio_fn(municipio, departamento, geo_idx)
        if mlat is not None:
            return {"lat": float(mlat), "lon": float(mlon), "fuente": "Centroide municipio (GADM)",
                    "estado": "[!] Revisar" if _es_urbana(localidad) else "[✓] OK",
                    "metodo": "no se halló el topónimo exacto; se usó el centro del municipio",
                    "detalle": (f"el sitio '{localidad}' no está en los gaceteros disponibles "
                                f"para {municipio}; coordenada provisional al centro del municipio")
                                if str(localidad).strip() and norm_geo(localidad) not in PALABRAS_VACIAS
                                else f"centro del municipio {municipio} ({minfo})"}

    # 4. CENTROIDE DE DEPARTAMENTO (Nivel 5)
    if nivel == 5 and centroide_depto_fn is not None:
        dlat, dlon = centroide_depto_fn(departamento, geo_idx)
        if dlat is not None:
            return {"lat": float(dlat), "lon": float(dlon), "fuente": "Centroide departamento (GADM)",
                    "estado": "[!] Revisar",
                    "metodo": "descripción general; se usó el centro del departamento",
                    "detalle": f"alta incertidumbre (escala departamental: {departamento})"}

    # 5. CENTROIDE DE PAÍS (Nivel 6)
    if nivel == 6:
        return {"lat": centroide_pais[0], "lon": centroide_pais[1], "fuente": "Centroide país",
                "estado": "[!] Revisar",
                "metodo": "solo se conoce el país; se usó el centro de Colombia",
                "detalle": "incertidumbre muy alta (escala nacional)"}

    # 6. REGLA DE ORO: nada confiable → no inventar
    return {"lat": None, "lon": None, "fuente": "", "estado": "[X] Error",
            "metodo": "no se pudo georreferenciar con las fuentes disponibles",
            "detalle": f"revisar '{localidad}' en {municipio}, {departamento}"}


def _es_urbana(localidad):
    """Heurística: ¿la localidad parece urbana (barrio/calle/sector)?"""
    l = norm_geo(localidad)
    return any(p in l for p in ("barrio", "calle", "carrera", "sector", "comuna",
                                "avenida", "manzana", "urbanizacion"))


def _promediar_si_cercanos(puntos, umbral_km=3.0):
    """
    Si varios puntos con el mismo nombre están MUY cerca (mismo sitio partido en
    varios registros), devuelve su centro promedio. Si están lejos (ambigüedad
    real: sitios distintos homónimos), devuelve None para no elegir a la ligera.
    """
    import math
    lats = puntos["lat"].astype(float).tolist()
    lons = puntos["lon"].astype(float).tolist()
    # distancia máxima entre puntos (aprox, en km)
    maxd = 0.0
    for i in range(len(lats)):
        for j in range(i + 1, len(lats)):
            dlat = (lats[i] - lats[j]) * 111
            dlon = (lons[i] - lons[j]) * 111 * math.cos(math.radians(lats[i]))
            maxd = max(maxd, math.hypot(dlat, dlon))
    if maxd <= umbral_km:
        return round(sum(lats) / len(lats), 6), round(sum(lons) / len(lons), 6)
    return None
