"""
geo_match.py — Emparejamiento robusto de municipios y departamentos contra GADM.
Protocolo SiB Colombia / Instituto Humboldt
© 2026 Ximena Bedoya Araque · Universidad CES · Colecciones Biológicas CBUCES

Una sola fuente de verdad para resolver el centroide de un municipio.
Soluciona dos errores graves:
  1. Homónimos: municipios con el mismo nombre en distintos departamentos
     (Barbosa, Caldas, Concordia...) → ahora la clave incluye el departamento.
  2. Diferencias de escritura (tildes, ñ, espacios): GADM concatena nombres
     ("PuertoLibertador", "Montelíbano", "NortedeSantander") → comparación robusta.

REGLA DE ORO: si no hay coincidencia CONFIABLE, NO se inventa una coordenada
cercana. Se devuelve None y un motivo, para que el curador lo revise.
"""

import re
import unicodedata
import difflib
from collections import defaultdict

import pandas as pd
import geopandas as gpd


# ───────────────────────── Normalización ─────────────────────────
def norm_geo(t):
    """Tildes fuera, ñ→n, puntuación/guiones → espacio, espacios colapsados, minúscula."""
    if t is None or (isinstance(t, float) and pd.isna(t)):
        return ""
    t = str(t).strip()
    if t.lower() in ("nan", ""):
        return ""
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn").lower()
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _compacto(s):
    """Sin espacios: hace que 'monte libano' == 'montelibano'."""
    return s.replace(" ", "")


# Umbral de similitud para coincidencia aproximada dentro del MISMO departamento.
UMBRAL_FUZZY = 0.88


# ───────────────────────── Construcción de índices ─────────────────────────
def construir_indices(ruta_gadm):
    """
    Lee GADM y construye los índices de búsqueda.
    Retorna dict con:
      muni:  depto_norm -> [ {mn, mc, var:[...], lat, lon, raw}, ... ]
      depto: depto_norm -> (lat, lon)        (centroide del departamento)
      depto_norms: set de departamentos normalizados (para resolver alias)
    Centroides calculados en EPSG:3116 (proyección Colombia) para precisión.
    """
    gdf = gpd.read_file(ruta_gadm)
    if gdf.crs is None:
        gdf = gdf.set_crs("EPSG:4326")
    gdf_p = gdf.to_crs("EPSG:3116")

    muni = defaultdict(list)
    for (_, r), (_, rp) in zip(gdf.iterrows(), gdf_p.iterrows()):
        c = gpd.GeoSeries([rp.geometry.centroid], crs="EPSG:3116").to_crs("EPSG:4326").iloc[0]
        dn = norm_geo(r.get("NAME_1"))
        mn = norm_geo(r.get("NAME_2"))
        var = [norm_geo(x) for x in str(r.get("VARNAME_2", "")).split("|")
               if x and str(x).lower() != "nan"]
        muni[dn].append({
            "mn": mn, "mc": _compacto(mn), "var": var,
            "lat": round(c.y, 6), "lon": round(c.x, 6), "raw": r.get("NAME_2"),
        })

    # Centroide de cada departamento (disolver polígonos por NAME_1)
    deptos_disuelto = gdf_p.dissolve(by="NAME_1").reset_index()
    depto = {}
    for _, r in deptos_disuelto.iterrows():
        c = gpd.GeoSeries([r.geometry.centroid], crs="EPSG:3116").to_crs("EPSG:4326").iloc[0]
        depto[norm_geo(r["NAME_1"])] = (round(c.y, 6), round(c.x, 6))

    return {"muni": dict(muni), "depto": depto, "depto_norms": set(depto.keys())}


# ───────────────────────── Resolución de departamento ─────────────────────────
def resolver_departamento(depto_raw, idx):
    """
    Devuelve el departamento normalizado que usa GADM, o None.
    Tolera 'Norte de Santander'->'NortedeSantander' y 'Guajira'->'La Guajira'.
    """
    dn = norm_geo(depto_raw)
    if not dn:
        return None
    if dn in idx["depto_norms"]:
        return dn
    dc = _compacto(dn)
    for g in idx["depto_norms"]:                       # compacto: Norte de Santander
        if _compacto(g) == dc:
            return g
    for g in idx["depto_norms"]:                       # subcadena: guajira ⊂ laguajira
        if dc in _compacto(g) or _compacto(g) in dc:
            return g
    return None


# ───────────────────────── Emparejamiento de municipio ─────────────────────────
def match_municipio(muni_raw, depto_raw, idx):
    """
    Busca el centroide del municipio dentro de SU departamento.
    Retorna (lat, lon, metodo) si hay coincidencia confiable,
    o (None, None, motivo) si NO la hay (para marcar revisión — nunca inventa).
    """
    dn = resolver_departamento(depto_raw, idx)
    if dn is None:
        return None, None, f"Departamento '{depto_raw}' no reconocido en GADM → revisar"

    cands = idx["muni"].get(dn, [])
    if not cands:
        return None, None, f"Departamento '{depto_raw}' sin municipios en GADM → revisar"

    mn = norm_geo(muni_raw)
    if not mn:
        return None, None, "Municipio vacío → revisar"
    mc = _compacto(mn)

    # 1. Coincidencia exacta (nombre o variante VARNAME_2)
    for c in cands:
        if mn == c["mn"] or mn in c["var"]:
            return c["lat"], c["lon"], "exacto"
    # 2. Coincidencia compacta (diferencias de espacios/tildes)
    for c in cands:
        if mc == c["mc"] or mc in [_compacto(v) for v in c["var"]]:
            return c["lat"], c["lon"], f"compacto ('{muni_raw}'≈'{c['raw']}')"
    # 3. Aproximada DENTRO del departamento (umbral alto)
    best, br = None, 0.0
    for c in cands:
        r = difflib.SequenceMatcher(None, mn, c["mn"]).ratio()
        if r > br:
            br, best = r, c
    if best and br >= UMBRAL_FUZZY:
        return best["lat"], best["lon"], f"aproximado {br:.0%} ('{muni_raw}'≈'{best['raw']}')"

    return None, None, (f"Municipio '{muni_raw}' no reconocido en {depto_raw} "
                        f"(más parecido: '{best['raw'] if best else '—'}' {br:.0%}) → revisar")


def centroide_departamento(depto_raw, idx):
    """Centroide del departamento (para Nivel 5). (lat, lon) o (None, None)."""
    dn = resolver_departamento(depto_raw, idx)
    return idx["depto"].get(dn, (None, None)) if dn else (None, None)
