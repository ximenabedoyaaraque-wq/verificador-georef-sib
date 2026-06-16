"""
Verificador de Georreferenciación — SiB Colombia
© 2026 Ximena Bedoya Araque · Universidad CES · Colecciones Biológicas CBUCES
Protocolo: Escobar D. et al. (2016) · Instituto Humboldt / ICN-UNAL
"""

import streamlit as st
import pandas as pd
import numpy as np
import os, sys, io

st.set_page_config(
    page_title="Verificador de Georreferenciación · SiB Colombia",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,300;0,9..144,400;0,9..144,500;1,9..144,300&family=DM+Sans:wght@300;400;500&family=DM+Mono:wght@400;500&display=swap');

:root {
  --ink:        #0D1F0F;
  --ink-2:      #2C3E2E;
  --ink-3:      #5A6B5C;
  --moss:       #2D5016;
  --fern:       #4A7C2F;
  --sage:       #8FAF72;
  --mist:       #EFF5E8;
  --paper:      #F9FAF7;
  --white:      #FFFFFF;
  --gold:       #C8953A;
  --warn:       #D97706;
  --danger:     #C0392B;
  --border:     #D4E0C8;
  --border-2:   #E8F0E1;
  --r:          8px;
  --r-lg:       14px;
  --mono:       'DM Mono', monospace;
  --sans:       'DM Sans', sans-serif;
  --serif:      'Fraunces', Georgia, serif;
  --shadow:     0 1px 3px rgba(13,31,15,.08), 0 4px 16px rgba(13,31,15,.04);
  --shadow-lg:  0 8px 32px rgba(13,31,15,.10);
}

@media (prefers-color-scheme: dark) {
  :root {
    --ink:      #E8F5E0;
    --ink-2:    #B8D4A8;
    --ink-3:    #7A9E6C;
    --moss:     #74C69D;
    --fern:     #52B788;
    --sage:     #3A7D50;
    --mist:     #132018;
    --paper:    #0D1A0F;
    --white:    #1A2C1C;
    --gold:     #F0B429;
    --warn:     #FBB034;
    --danger:   #E74C3C;
    --border:   #2A4030;
    --border-2: #1E3025;
    --shadow:   0 1px 3px rgba(0,0,0,.3);
    --shadow-lg:0 8px 32px rgba(0,0,0,.4);
  }
}

html, body, [class*="css"] { font-family: var(--sans) !important; color: var(--ink) !important; }

section[data-testid="stSidebar"] { background: var(--paper) !important; border-right: 1px solid var(--border) !important; }
section[data-testid="stSidebar"] > div { padding: 28px 20px 20px !important; }
.main .block-container { padding: 40px 48px 48px !important; max-width: 1280px !important; }

.logo-wrap { margin-bottom: 24px; }
.logo-name { font-family: var(--serif) !important; font-size: 21px; font-weight: 500; color: var(--moss) !important; line-height: 1.2; letter-spacing: -.01em; }
.logo-sub  { font-size: 16px; font-weight: 400; color: var(--ink-3) !important; text-transform: uppercase; letter-spacing: .12em; margin-top: 3px; }

.upload-label { font-size: 16px; font-weight: 500; color: var(--ink-3) !important; text-transform: uppercase; letter-spacing: .1em; margin-bottom: 5px; margin-top: 14px; display: block; }

.pill { display: inline-flex; align-items: center; gap: 5px; padding: 4px 10px; border-radius: 20px; font-size: 15px; font-weight: 500; margin-bottom: 4px; width: 100%; }
.pill-ok   { background: #E8F5E0; color: #2D5016; }
.pill-warn { background: #FEF3C7; color: #7C5A00; }
.pill-err  { background: #FEE2E2; color: #7F1D1D; }
.pill-nd   { background: var(--mist); color: var(--ink-3); }

.copy { font-size: 15px; color: var(--ink-3) !important; line-height: 1.7; margin-top: 20px; padding-top: 16px; border-top: 1px solid var(--border-2); }

.hero { margin-bottom: 40px; }
.hero-eyebrow { font-size: 16px; font-weight: 500; color: var(--fern) !important; text-transform: uppercase; letter-spacing: .16em; margin-bottom: 10px; }
.hero-title { font-family: var(--serif) !important; font-size: clamp(34px, 5vw, 52px); font-weight: 300; color: var(--ink) !important; line-height: 1.15; letter-spacing: -.02em; margin-bottom: 14px; }
.hero-title em { font-style: italic; color: var(--fern) !important; }
.hero-desc { font-size: 16px; color: var(--ink-3) !important; max-width: 560px; line-height: 1.65; }

.metric-grid { display: flex; gap: 12px; margin: 28px 0 36px; }
.metric-card { background: var(--white); border: 1px solid var(--border); border-radius: var(--r-lg); padding: 18px 20px; flex: 1; box-shadow: var(--shadow); position: relative; overflow: hidden; }
.metric-card::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, var(--fern), var(--sage)); }
.metric-num { font-family: var(--serif) !important; font-size: 36px; font-weight: 400; color: var(--moss) !important; line-height: 1; margin-bottom: 5px; }
.metric-label { font-size: 16px; font-weight: 500; color: var(--ink-3) !important; text-transform: uppercase; letter-spacing: .1em; }

.steps { margin: 24px 0; }
.step-item { display: flex; align-items: flex-start; gap: 14px; margin-bottom: 12px; }
.step-num { width: 24px; height: 24px; border-radius: 50%; background: var(--mist); border: 1.5px solid var(--border); display: flex; align-items: center; justify-content: center; font-size: 15px; font-weight: 500; color: var(--fern) !important; flex-shrink: 0; margin-top: 1px; }
.step-text { font-size: 15px; color: var(--ink-2) !important; line-height: 1.5; padding-top: 2px; }

.section-head { display: flex; align-items: center; gap: 10px; margin: 32px 0 16px; }
.section-line { flex: 1; height: 1px; background: var(--border-2); }
.section-label { font-size: 16px; font-weight: 500; color: var(--ink-3) !important; text-transform: uppercase; letter-spacing: .12em; white-space: nowrap; }

.result-bar { display: flex; gap: 10px; margin: 20px 0; flex-wrap: wrap; }
.result-chip { display: flex; align-items: center; gap: 7px; background: var(--white); border: 1px solid var(--border); border-radius: var(--r); padding: 10px 16px; box-shadow: var(--shadow); }
.chip-dot { width: 8px; height: 8px; border-radius: 50%; }
.chip-num { font-family: var(--serif) !important; font-size: 23px; font-weight: 400; color: var(--ink) !important; line-height: 1; }
.chip-label { font-size: 16px; color: var(--ink-3) !important; text-transform: uppercase; letter-spacing: .08em; }

button[data-baseweb="tab"] { font-family: var(--sans) !important; font-size: 16px !important; font-weight: 500 !important; letter-spacing: .03em !important; text-transform: uppercase !important; }

.info-box { background: var(--mist); border: 1px solid var(--border); border-left: 3px solid var(--fern); border-radius: var(--r); padding: 14px 16px; font-size: 15px; color: var(--ink-2) !important; line-height: 1.6; }

.dl-card { background: var(--white); border: 1px solid var(--border); border-radius: var(--r-lg); padding: 28px 32px; box-shadow: var(--shadow-lg); margin-bottom: 24px; }
.dl-title { font-family: var(--serif) !important; font-size: 23px; font-weight: 400; color: var(--ink) !important; margin-bottom: 8px; }
.dl-desc { font-size: 15px; color: var(--ink-3) !important; line-height: 1.6; margin-bottom: 20px; }
.dl-list { font-size: 16px; color: var(--ink-2) !important; line-height: 1.8; padding-left: 0; list-style: none; }
.dl-list li::before { content: "→  "; color: var(--fern); font-weight: 500; }

.reference { font-size: 15px; color: var(--ink-3) !important; line-height: 1.7; padding: 16px; background: var(--mist); border-radius: var(--r); border: 1px solid var(--border-2); font-style: italic; }
</style>
""", unsafe_allow_html=True)

# ── Rutas ──────────────────────────────────────────
BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
sys.path.insert(0, os.path.join(BASE_DIR, "bloques"))

_posibles_gadm = [
    os.path.join(BASE_DIR, "datos", "gadm41_COL_2.json"),
    "/mount/src/verificador-georef-sib/datos/gadm41_COL_2.json",
    "datos/gadm41_COL_2.json",
]
GADM_PATH = next((p for p in _posibles_gadm if os.path.exists(p)), None)

# Gacetero de nombres geográficos (IGAC) para la cascada del manual
_posibles_gacetero = [
    os.path.join(BASE_DIR, "datos", "gacetero_igac_limpio.parquet"),
    "/mount/src/verificador-georef-sib/datos/gacetero_igac_limpio.parquet",
    "datos/gacetero_igac_limpio.parquet",
]
GACETERO_PATH = next((p for p in _posibles_gacetero if os.path.exists(p)), None)

# ── Helper: normalizar valor de validación ─────────
def _val_color(val: str) -> str:
    """Devuelve color hex según el texto de validación (emojis O texto corchete)."""
    v = str(val).strip()
    if "OK" in v or "✅" in v:      return "#4A7C2F"
    if "Revisar" in v or "⚠" in v:  return "#D97706"
    if "Error" in v or "❌" in v:   return "#C0392B"
    return "#9CA3AF"

def _val_desc(val: str) -> str:
    v = str(val).strip()
    if "OK" in v or "✅" in v:      return "Dentro del municipio reportado"
    if "Revisar" in v or "⚠" in v:  return "Municipio vecino — revisar"
    if "Error" in v or "❌" in v:   return "Fuera de Colombia o municipio incorrecto"
    return "Sin validación espacial"

def _post_procesar_universal(df, gadm_path=None, gacetero_path=None):
    """Incertidumbre, cascada de georreferenciación (manual SiB) y comentarios."""
    import re as _re, unicodedata as _ud

    df = df.copy().reset_index(drop=True)  # evita desalineación de índices

    def _norm(t):
        if pd.isna(t): return ""
        t = _ud.normalize("NFD", str(t).strip())
        return "".join(c for c in t if _ud.category(c) != "Mn").lower()

    # 1. Incertidumbre de coordenadas
    def _calc_inc(row):
        datum = str(row.get("Datum", "")).strip()
        fmt   = str(row.get("formato_coordenada", "")).strip()
        if not fmt:
            s = str(row.get("Latitud original", "")).strip().replace(".0", "")
            if "''" in s:                        fmt = "GMS"
            elif "°" in s:                       fmt = "GMD"
            elif _re.match(r"^-?\d{5,10}$", s): fmt = "entero sin punto"
            else:
                try:    float(s); fmt = "decimal"
                except: fmt = ""
        inc_d = 500 if datum in ("", "nan", "WGS 84 (asumido)") else 0
        inc_c = {"GMS": 44, "GMD": 262, "entero sin punto": 2, "decimal": 157}.get(fmt, 0)
        total = inc_d + inc_c
        return int(total) if total > 0 else None

    df["Incertidumbre de coordenadas (m)"] = df.apply(_calc_inc, axis=1)

    # 2. CASCADA DE GEORREFERENCIACIÓN (manual SiB): topónimo → municipio → depto → país → marca
    for col in ("Latitud georreferenciada", "Longitud georreferenciada"):
        if col not in df.columns:
            df[col] = ""
    df["Latitud georreferenciada"]  = df["Latitud georreferenciada"].astype(object)
    df["Longitud georreferenciada"] = df["Longitud georreferenciada"].astype(object)
    for col in ("Método de georreferenciación", "Fuentes de georreferenciación"):
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(object)

    # FIX coordenadas originales: limpiar caracteres inválidos (?, coma→punto, etc.)
    # Manual p.67: conservar la coordenada original del colector; un '?' o coma son errores
    # de digitación evidentes que deben corregirse antes de convertir.
    import re as _re, unicodedata as _ud
    def _limpiar_coord(val):
        """Quita caracteres inválidos (?, *, espacios extra) y normaliza separador decimal."""
        if pd.isna(val) or str(val).strip() in ("", "nan"): return val
        s = str(val).strip()
        s = _re.sub(r"[?*°\s]", lambda m: " " if m.group() == " " else "", s)
        s = s.replace(",", ".")
        s = _re.sub(r"\s+", " ", s).strip()
        return s if s else val
    for c_orig in ("Latitud original", "Longitud original"):
        if c_orig in df.columns:
            df[c_orig] = df[c_orig].apply(_limpiar_coord)

    # FIX normalizador para comparación de nombres (anti-PuertoLibertador/Puerto Libertador)
    def _norm_muni(t):
        if pd.isna(t) or str(t).strip() in ("", "nan"): return ""
        t = _ud.normalize("NFD", str(t).strip())
        t = "".join(c for c in t if _ud.category(c) != "Mn").lower()
        return _re.sub(r"\s+", "", _re.sub(r"[^a-z0-9]+", " ", t))  # compacto sin espacios

    # Cargar cascada, gaceteros e índice GADM (una sola vez)
    idx_geo = None
    gaceteros = []
    _cg = None
    if gadm_path and os.path.exists(gadm_path):
        try:
            import geo_match as _gm
            import cascada_gaceteros as _cg
            idx_geo = _gm.construir_indices(gadm_path)
            if gacetero_path and os.path.exists(gacetero_path):
                gaceteros.append(("IGAC", _cg.cargar_gacetero(gacetero_path)))
        except Exception as _e:
            idx_geo = None
            _cg = None

    def _sin_coord(i):
        v = df.at[i, "Latitud georreferenciada"]
        return pd.isna(v) or str(v).strip() in ("", "nan")

    # FIX coordenada del mismo topónimo (Manual Nivel 3 Caso 1, p.74):
    # Si hay registros del mismo sitio (misma localidad+municipio) con coordenada
    # verificada [✓] OK, usar su promedio como referencia — más fiel al manual que
    # el centroide del municipio (que corresponde al Nivel 4, no al 3).
    import math as _math
    col_val_esp = ("Resultado validación espacial" if "Resultado validación espacial" in df.columns
                   else "validacion_b2" if "validacion_b2" in df.columns else "")
    _ref_toponimo = {}  # (loc_norm, muni_norm) -> (lat_prom, lon_prom, radio_m, n)
    if col_val_esp:
        _col_loc = "*Localidad estandarizada" if "*Localidad estandarizada" in df.columns else ""
        if _col_loc:
            for (loc_k, muni_k), grp in df.groupby([_col_loc, "*Municipio"]):
                ok = grp[grp[col_val_esp].astype(str).str.contains("OK|✅", na=False)]
                ok = ok[ok["Latitud georreferenciada"].notna()]
                if len(ok) >= 2:
                    lats = ok["Latitud georreferenciada"].astype(float).tolist()
                    lons = ok["Longitud georreferenciada"].astype(float).tolist()
                    lp, lonp = sum(lats)/len(lats), sum(lons)/len(lons)
                    radio = max(
                        _math.hypot((lp-la)*111, (lonp-lo)*111*_math.cos(_math.radians(lp)))*1000
                        for la, lo in zip(lats, lons)) if len(lats) > 1 else 0
                    key = (_norm_muni(str(loc_k)), _norm_muni(str(muni_k)))
                    _ref_toponimo[key] = (round(lp,6), round(lonp,6), round(radio), len(ok))

    # Inicializar columnas temporales de la cascada
    df["_cascada_estado"]  = ""
    df["_cascada_detalle"] = ""
    df["_muni_gadm"]       = ""   # municipio real según GADM (para comentarios)

    if idx_geo is not None and _cg is not None:
        for idx in df.index:
            nivel = int(df.at[idx, "Nivel_final"]) if pd.notna(df.at[idx, "Nivel_final"]) else None
            loc   = df.at[idx, "*Localidad estandarizada"] if "*Localidad estandarizada" in df.columns else ""
            muni  = str(df.at[idx, "*Municipio"])    if "*Municipio"    in df.columns else ""
            depto = str(df.at[idx, "*Departamento"]) if "*Departamento" in df.columns else ""

            if nivel == 7:
                df.at[idx, "Método de georreferenciación"] = "No se georreferencia (Nivel 7)"
                continue

            # Nivel 1 con coordenada: verificar si es Caso 3 (cae en otro municipio)
            if nivel == 1 and not _sin_coord(idx):
                lat_g = float(df.at[idx, "Latitud georreferenciada"])
                lon_g = float(df.at[idx, "Longitud georreferenciada"])
                val_actual = str(df.at[idx, col_val_esp]).strip() if col_val_esp else ""
                muni_det = str(df.at[idx, "municipio_detectado"]).strip() if "municipio_detectado" in df.columns else ""

                # FIX Caso 3: si la coord cae fuera del municipio reportado,
                # reasignar al centroide del municipio y documentarlo (Manual Tabla 10 Caso 3)
                if "[X]" in val_actual or "Error" in val_actual:
                    c3lat, c3lon, c3info = _gm.match_municipio(muni, depto, idx_geo)
                    if c3lat is not None:
                        df.at[idx, "Latitud georreferenciada"]  = c3lat
                        df.at[idx, "Longitud georreferenciada"] = c3lon
                        df.at[idx, "Método de georreferenciación"] = f"Reasignado a centroide de {muni} (Manual Tabla 10, Caso 3)"
                        df.at[idx, "_cascada_estado"]  = "[X] Error — reasignado"
                        df.at[idx, "_muni_gadm"]       = muni_det
                        df.at[idx, "_cascada_detalle"] = (
                            f"COORDENADA ORIGINAL FUERA DEL MUNICIPIO REPORTADO: "
                            f"la coordenada ({lat_g:.5f}, {lon_g:.5f}) cae en '{muni_det or 'otro municipio'}', "
                            f"no en {muni} ({depto}). "
                            f"Según el manual (Tabla 10, Caso 3), cuando la coordenada no corresponde "
                            f"al lugar descrito no se valida: se reasignó al centroide de {muni}. "
                            f"Verificar si el error está en la coordenada original o en el municipio reportado.")
                    continue

                # FIX falso positivo escritura GADM: PuertoLibertador ≠ Puerto Libertador
                # GADM concatena nombres compuestos; el colector los escribe con espacios.
                # Si el normalizador compacto los iguala → es el mismo municipio → OK real.
                if ("[!]" in val_actual or "Revisar" in val_actual) and muni_det:
                    if _norm_muni(muni_det) == _norm_muni(muni):
                        # Mismo municipio, solo diferencia tipográfica GADM vs colector
                        if col_val_esp:
                            df.at[idx, col_val_esp] = "[✓] OK"
                        df.at[idx, "_cascada_estado"]  = "[✓] OK"
                        df.at[idx, "_cascada_detalle"] = (
                            f"Coordenada validada dentro de {muni}. "
                            f"Nota: GADM escribe este municipio como '{muni_det}' (sin espacios), "
                            f"convención tipográfica distinta a la del colector ('{muni}') — "
                            f"es el mismo lugar, no hay error en los datos.")
                    elif _norm_muni(muni_det) and _norm_muni(muni_det) != _norm_muni(muni):
                        # Zona limítrofe: coordenada en municipio vecino según GADM
                        df.at[idx, "_cascada_estado"]  = "[!] Revisar"
                        df.at[idx, "_cascada_detalle"] = (
                            f"Zona limítrofe: la coordenada cae en '{muni_det}' según GADM, "
                            f"pero el colector reportó {muni}. "
                            f"En zonas de límite municipal, el conocimiento de campo del colector "
                            f"prevalece sobre el límite cartográfico (Manual Tabla 10 Caso 2). "
                            f"Se mantiene la coordenada original y el municipio reportado. "
                            f"Verificar con el colector si hubo imprecisión en el límite.")
                    continue

                # Nivel 1 OK normal
                if str(df.at[idx, "Método de georreferenciación"]).strip() in ("", "nan"):
                    df.at[idx, "Método de georreferenciación"] = "Coordenada original del colector"
                continue

            # Niveles 2-6 sin coordenada: intentar con referencia del mismo topónimo primero
            loc_key = (_norm_muni(str(loc)), _norm_muni(muni))
            if loc_key in _ref_toponimo and str(loc).strip().lower() not in ("sin datos","nan",""):
                lp, lonp, radio, n_ref = _ref_toponimo[loc_key]
                df.at[idx, "Latitud georreferenciada"]  = lp
                df.at[idx, "Longitud georreferenciada"] = lonp
                df.at[idx, "Método de georreferenciación"] = (
                    f"Coordenada de referencia del mismo sitio (promedio de {n_ref} registros "
                    f"verificados del mismo topónimo en {muni})")
                df.at[idx, "Fuentes de georreferenciación"] = "Registros verificados del mismo lote"
                df.at[idx, "_cascada_estado"]  = "[✓] OK"
                df.at[idx, "_cascada_detalle"] = (
                    f"No hay coordenada original para este registro, pero existen {n_ref} registros "
                    f"verificados [✓] del mismo sitio ('{loc}', {muni}). "
                    f"Según el manual (Nivel 3 Caso 1, p.74), se usa su coordenada promedio "
                    f"como referencia cartográfica. Incertidumbre = dispersión del lote: {radio} m.")
                continue

            # Cascada normal: gacetero → municipio → depto → país → marca
            res = _cg.georreferenciar_localidad(
                loc, muni, depto, nivel, gaceteros, idx_geo,
                _gm.match_municipio, _gm.centroide_departamento)

            if res["lat"] is not None:
                df.at[idx, "Latitud georreferenciada"]  = res["lat"]
                df.at[idx, "Longitud georreferenciada"] = res["lon"]
            df.at[idx, "Método de georreferenciación"] = res["metodo"]
            if res["fuente"]:
                df.at[idx, "Fuentes de georreferenciación"] = res["fuente"]
            df.at[idx, "_cascada_estado"]  = res["estado"]
            df.at[idx, "_cascada_detalle"] = res["detalle"]

    # 3. Comentarios detallados por resultado de validación
    _col_val = (
        "Resultado validación espacial" if "Resultado validación espacial" in df.columns
        else "validacion_b2" if "validacion_b2" in df.columns
        else ""
    )

    def _comentario(row):
        com      = str(row.get("Comentarios de la georreferenciación", "")).strip()
        nivel    = int(row.get("Nivel_final", 0) or 0)
        val      = str(row.get(_col_val, "")).strip() if _col_val else ""
        muni_det = str(row.get("municipio_detectado", "")).strip()
        lat_o    = str(row.get("Latitud original", "")).strip()
        fmt      = str(row.get("formato_coordenada", "")).strip()
        metodo_c  = str(row.get("Método de georreferenciación", "")).strip()
        detalle_c = str(row.get("_cascada_detalle", "")).strip()

        if nivel == 7:
            return ("Nivel 7: información dudosa o contradictoria. Según el manual SiB, "
                    "los registros de Nivel 7 no se georreferencian. "
                    "Requiere revisión del curador para resolver la inconsistencia.")

        # Usar el comentario amigable de la cascada si existe
        if detalle_c and detalle_c.lower() not in ("nan", ""):
            return detalle_c

        # Fallback a comentarios por estado de validación
        if "OK" in val or "✅" in val:
            suf = ("[✓] OK — coordenada validada dentro del municipio reportado.")
        elif "Revisar" in val or "⚠" in val or "[!]" in val:
            suf = (f"[!] REQUIERE REVISIÓN — la coordenada cae en '{muni_det}', "
                   f"distinto al municipio reportado. Verificar: error de digitación, "
                   f"municipio incorrecto, o localidad en zona limítrofe.")
        elif "Error" in val or "❌" in val or "[X]" in val:
            suf = (f"[X] ERROR — la coordenada original no corresponde al municipio reportado. "
                   f"Valor original: {lat_o} (formato: {fmt}). "
                   f"Municipio detectado por GADM: '{muni_det}'. "
                   f"Se reasignó al centroide del municipio reportado (Manual Tabla 10 Caso 3). "
                   f"Requiere verificación manual.")
        else:
            return com if com else f"Nivel {nivel}: {metodo_c}."

        sep = " | " if com and com.lower() != "nan" else ""
        base = com if com and com.lower() != "nan" else ""
        return base + sep + suf

    df["Comentarios de la georreferenciación"] = df.apply(_comentario, axis=1)
    for c in ("_cascada_estado", "_cascada_detalle", "_muni_gadm"):
        if c in df.columns:
            df = df.drop(columns=[c])
    return df

# ── Sidebar ────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="logo-wrap">
      <div class="logo-name">Verificador de<br>Georreferenciación</div>
      <div class="logo-sub">SiB Colombia · Instituto Humboldt</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<span class="upload-label">Base con coordenadas (.xlsx)</span>', unsafe_allow_html=True)
    file_180 = st.file_uploader("", type=["xlsx","xls"], key="file_180", label_visibility="collapsed")

    st.markdown('<span class="upload-label">Base sin coordenadas (.xlsx)</span>', unsafe_allow_html=True)
    file_84  = st.file_uploader("", type=["xlsx","xls"], key="file_84",  label_visibility="collapsed")

    st.markdown("<br>", unsafe_allow_html=True)

    gadm_ok = GADM_PATH is not None and os.path.exists(GADM_PATH)
    if gadm_ok:
        st.markdown('<div class="pill pill-ok">✓ &nbsp;Capa GADM Colombia activa</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="pill pill-warn">⚠ &nbsp;GADM no encontrado — validación espacial limitada</div>', unsafe_allow_html=True)

    ejecutar = st.button(
        "▶  Ejecutar análisis",
        use_container_width=True,
        type="primary",
        disabled=(file_180 is None or file_84 is None),
    )

    st.markdown("""
    <div class="copy">
      © 2026 Ximena Bedoya Araque<br>
      Estudiante de Ecología · Universidad CES<br>
      Pasantía en Colecciones Biológicas CBUCES<br>
      Medellín, Colombia<br><br>
      Basado en Escobar D. et al. (2016)<br>
      Instituto Humboldt – ICN/UNAL<br>
      Licencia CC BY-NC 4.0
    </div>
    """, unsafe_allow_html=True)

# ── Estado de sesión ───────────────────────────────
if "df_resultado" not in st.session_state:
    st.session_state.df_resultado  = None
    st.session_state.excel_bytes   = None
    st.session_state.procesado     = False

# ── Procesamiento ──────────────────────────────────
if ejecutar and file_180 and file_84:
    with st.spinner("Procesando registros…"):
        try:
            from verificador_georef_completo_6 import (
                aplicar_bloque1, aplicar_bloque3, aplicar_bloque5,
                aplicar_bloque6, aplicar_bloque7, aplicar_bloque8,
                aplicar_bloque9, aplicar_bloque10,
            )

            with open("/tmp/base_180.xlsx", "wb") as f: f.write(file_180.read())
            with open("/tmp/base_84.xlsx",  "wb") as f: f.write(file_84.read())

            prog = st.progress(0,  text="Estandarizando coordenadas…")
            df   = aplicar_bloque1("/tmp/base_84.xlsx", "/tmp/base_180.xlsx")

            prog.progress(20, text="Verificando campos Darwin Core…")
            df, _ = aplicar_bloque5(df)

            prog.progress(28, text="Comparando localidad original vs estandarizada…")
            df = aplicar_bloque6(df)

            prog.progress(35, text="Clasificando niveles de calidad…")
            df = aplicar_bloque7(df)

            prog.progress(50, text="Validando coordenadas contra municipios GADM…")
            if gadm_ok and GADM_PATH:
                df = aplicar_bloque8(df, GADM_PATH)
            else:
                df["Nivel_final"]         = df["Nivel_inicial"].copy()
                df["lat_wgs84"]           = df["lat_decimal_calculada"]
                df["lon_wgs84"]           = df["lon_decimal_calculada"]
                df["validacion_b2"]       = df["conversion_estado"].map(
                    {"OK":"✅ OK","Revisar":"⚠ Revisar",
                     "Error":"❌ Error","sin coordenadas":""}
                ).fillna("")
                df["municipio_detectado"] = ""
                df["depto_detectado"]     = ""
                df["mensaje_b2"]          = ""

            prog.progress(70, text="Asignando centroides…")
            if gadm_ok and GADM_PATH:
                df = aplicar_bloque9(df, GADM_PATH, usar_nominatim=False)

            prog.progress(78, text="Verificando elevación…")
            df = aplicar_bloque3(df)

            # Sacar elevación API de columnas internas → columnas visibles en Excel
            if "elevacion_api" in df.columns:
                df["Elevación API (msnm)"]  = df["elevacion_api"]
                df["Validación elevación"]   = df["elevacion_estado"].map(
                    {"OK": "[✓] OK", "Revisar": "[!] Revisar"}).fillna("")
                df["Nota elevación"]         = df["elevacion_nota"]

            prog.progress(85, text="Generando reporte Excel…")
            st.session_state.excel_bytes = aplicar_bloque10(df, idioma=None)

            prog.progress(90, text="Post-procesamiento: cascada del manual (gacetero→municipio) y comentarios…")
            df = _post_procesar_universal(df, GADM_PATH, GACETERO_PATH)
            st.session_state.excel_bytes = aplicar_bloque10(df, idioma=None)

            st.session_state.df_resultado = df
            st.session_state.procesado    = True
            prog.progress(100, text="¡Listo!")

        except Exception as e:
            st.error(f"Error durante el procesamiento: {e}")
            import traceback
            st.code(traceback.format_exc())

# ── Pantalla de bienvenida ─────────────────────────
if not st.session_state.procesado:
    st.markdown("""
    <div class="hero">
      <div class="hero-eyebrow">Protocolo SiB Colombia · Instituto Humboldt</div>
      <div class="hero-title">Verificador de<br><em>Georreferenciación</em></div>
      <div class="hero-desc">
        Herramienta para la validación y georreferenciación de localidades
        en colecciones biológicas, basada en el protocolo del Instituto de
        Investigación de Recursos Biológicos Alexander von Humboldt e
        Instituto de Ciencias Naturales (UNAL).
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="metric-grid">
      <div class="metric-card"><div class="metric-num">10</div><div class="metric-label">Procesos automatizados</div></div>
      <div class="metric-card"><div class="metric-num">7</div><div class="metric-label">Niveles de calidad</div></div>
      <div class="metric-card"><div class="metric-num">50k+</div><div class="metric-label">Registros soportados</div></div>
      <div class="metric-card"><div class="metric-num">264</div><div class="metric-label">Registros de prueba</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="section-head">
      <div class="section-label">Cómo usar</div>
      <div class="section-line"></div>
    </div>
    <div class="steps">
      <div class="step-item"><div class="step-num">1</div><div class="step-text">Sube tu base <b>con coordenadas</b> (.xlsx) desde el panel izquierdo</div></div>
      <div class="step-item"><div class="step-num">2</div><div class="step-text">Sube tu base <b>sin coordenadas</b> (.xlsx)</div></div>
      <div class="step-item"><div class="step-num">3</div><div class="step-text">Presiona <b>▶ Ejecutar análisis</b> y espera el procesamiento</div></div>
      <div class="step-item"><div class="step-num">4</div><div class="step-text">Revisa los resultados en el <b>visor de puntos</b> y la tabla</div></div>
      <div class="step-item"><div class="step-num">5</div><div class="step-text">Descarga el <b>reporte Excel</b> con colores, niveles y comentarios</div></div>
    </div>
    <div class="info-box">
      💡 Ambas bases deben venir estandarizadas del proceso previo
      (Verificador de Localidades — Taller 1), con la columna
      <b>*Localidad estandarizada</b> ya corregida.
    </div>
    """, unsafe_allow_html=True)

# ── Resultados ─────────────────────────────────────
else:
    df    = st.session_state.df_resultado
    total = len(df)
    n1    = int((df["Nivel_final"] == 1).sum())
    n2_6  = int(df["Nivel_final"].isin([2,3,4,5,6]).sum())
    n7    = int((df["Nivel_final"] == 7).sum())

    col_val = (
        "Resultado validación espacial" if "Resultado validación espacial" in df.columns
        else "validacion_b2" if "validacion_b2" in df.columns
        else ""
    )

    # FIX: str.contains para soportar emojis Y texto corchete
    if col_val:
        val_ok  = int(df[col_val].str.contains("OK|✅",     na=False).sum())
        val_rev = int(df[col_val].str.contains("Revisar|⚠", na=False).sum())
        val_err = int(df[col_val].str.contains("Error|❌",  na=False).sum())
    else:
        val_ok = val_rev = val_err = 0

    st.markdown(f"""
    <div class="result-bar">
      <div class="result-chip">
        <div class="chip-dot" style="background:#4A7C2F"></div>
        <div><div class="chip-num">{total}</div><div class="chip-label">Total</div></div>
      </div>
      <div class="result-chip">
        <div class="chip-dot" style="background:#4A7C2F"></div>
        <div><div class="chip-num">{n1}</div><div class="chip-label">Con coordenadas</div></div>
      </div>
      <div class="result-chip">
        <div class="chip-dot" style="background:#8FAF72"></div>
        <div><div class="chip-num">{n2_6}</div><div class="chip-label">Georreferenciados</div></div>
      </div>
      <div class="result-chip">
        <div class="chip-dot" style="background:#999"></div>
        <div><div class="chip-num">{n7}</div><div class="chip-label">Nivel 7</div></div>
      </div>
      <div class="result-chip">
        <div class="chip-dot" style="background:#4A7C2F"></div>
        <div><div class="chip-num">{val_ok}</div><div class="chip-label">✅ OK</div></div>
      </div>
      <div class="result-chip">
        <div class="chip-dot" style="background:#D97706"></div>
        <div><div class="chip-num">{val_rev}</div><div class="chip-label">⚠ Revisar</div></div>
      </div>
      <div class="result-chip">
        <div class="chip-dot" style="background:#C0392B"></div>
        <div><div class="chip-num">{val_err}</div><div class="chip-label">❌ Error</div></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["🗺️  Visor de puntos", "📋  Tabla de resultados", "⬇️  Descargar reporte"])

    # ── Pestaña 1: Mapa ───────────────────────────
    with tab1:
        try:
            import folium
            from streamlit_folium import st_folium
            from folium.plugins import MarkerCluster

            df_m = df.copy()

            m = folium.Map(location=[5.5,-74.5], zoom_start=6,
                           tiles="CartoDB Positron", prefer_canvas=True)
            folium.TileLayer(
                "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
                attr="Esri", name="Satelital", overlay=False, control=True
            ).add_to(m)

            # ── Capa del gacetero: límites municipales (GADM) ──
            # Contorno de los 1119 municipios + nombre al acercar el zoom.
            _muni_path = next((p for p in [
                os.path.join(BASE_DIR, "datos", "municipios_simpl.geojson"),
                "/mount/src/verificador-georef-sib/datos/municipios_simpl.geojson",
                "datos/municipios_simpl.geojson",
            ] if os.path.exists(p)), None)
            if _muni_path:
                _gj = folium.GeoJson(
                    _muni_path,
                    name="🗺️ Límites municipales (gacetero)",
                    style_function=lambda f: {
                        "fillColor": "#000000", "fillOpacity": 0,
                        "color": "#4A7C2F", "weight": 0.6, "opacity": 0.55,
                    },
                    highlight_function=lambda f: {
                        "weight": 2, "color": "#2D5016", "fillOpacity": 0.08,
                    },
                    tooltip=folium.GeoJsonTooltip(
                        fields=["muni", "depto"],
                        aliases=["Municipio:", "Departamento:"],
                        sticky=True,
                    ),
                    show=True,
                )
                _gj.add_to(m)
                # Etiquetas de nombre SOLO en municipios que tienen registros
                # (evita encimar 1119 nombres ilegibles; el contorno sí es de todo el país).
                try:
                    import geopandas as _gpd
                    _gmuni = _gpd.read_file(_muni_path)
                    _munis_con_datos = set()
                    if "*Municipio" in df_m.columns:
                        import unicodedata as _u, re as _r
                        def _nz(t):
                            t = _u.normalize("NFD", str(t).strip())
                            t = "".join(c for c in t if _u.category(c) != "Mn").lower()
                            return _r.sub(r"\s+", "", _r.sub(r"[^a-z0-9]+", " ", t))
                        _munis_con_datos = {_nz(x) for x in df_m["*Municipio"].dropna().unique()}
                        _gmuni["_k"] = _gmuni["muni"].map(_nz)
                        _gmuni = _gmuni[_gmuni["_k"].isin(_munis_con_datos)]
                    _label_layer = folium.FeatureGroup(name="🔤 Nombres de municipio", show=True)
                    _gmuni_p = _gmuni.to_crs("EPSG:3116")
                    for (_, r), (_, rp) in zip(_gmuni.iterrows(), _gmuni_p.iterrows()):
                        c = _gpd.GeoSeries([rp.geometry.centroid], crs="EPSG:3116").to_crs("EPSG:4326").iloc[0]
                        folium.map.Marker(
                            [c.y, c.x],
                            icon=folium.DivIcon(html=(
                                f'<div style="font-family:DM Sans,sans-serif;font-size:11px;'
                                f'font-weight:600;color:#2D5016;white-space:nowrap;'
                                f'text-shadow:0 0 3px #fff,0 0 3px #fff,0 0 3px #fff;'
                                f'transform:translate(-50%,-50%)">{r["muni"]}</div>'))
                        ).add_to(_label_layer)
                    _label_layer.add_to(m)
                except Exception:
                    pass

            # ── Clusters coloreados por ESTADO (no por cantidad) ──
            # Así el círculo grande tiene el mismo color que los puntos pequeños adentro.
            def _cluster_icon(hexcolor):
                return (
                    "function(cluster){"
                    "var n=cluster.getChildCount();"
                    "return new L.DivIcon({"
                    "html:'<div style=\"background:%s;opacity:.85;width:38px;height:38px;"
                    "border-radius:50%%;display:flex;align-items:center;justify-content:center;"
                    "border:2px solid white;box-shadow:0 1px 4px rgba(0,0,0,.3)\">"
                    "<span style=\"color:white;font-weight:700;font-size:13px;"
                    "font-family:sans-serif\">'+n+'</span></div>',"
                    "className:'mcc',iconSize:new L.Point(38,38)});}" % hexcolor
                )
            _opts = {"maxClusterRadius":40,"disableClusteringAtZoom":12}
            clusters = {
                "OK":      MarkerCluster(name="🟢 OK",              options=_opts, icon_create_function=_cluster_icon("#4A7C2F")).add_to(m),
                "Revisar": MarkerCluster(name="🟠 Revisar",         options=_opts, icon_create_function=_cluster_icon("#D97706")).add_to(m),
                "Error":   MarkerCluster(name="🔴 Error",           options=_opts, icon_create_function=_cluster_icon("#C0392B")).add_to(m),
                "Sin":     MarkerCluster(name="⚪ Sin validación",   options=_opts, icon_create_function=_cluster_icon("#9CA3AF")).add_to(m),
            }

            col_lat = next((c for c in ["Latitud georreferenciada", "lat_wgs84", "lat_decimal_calculada"] if c in df_m.columns), "lat_decimal_calculada")
            col_lon = next((c for c in ["Longitud georreferenciada", "lon_wgs84", "lon_decimal_calculada"] if c in df_m.columns), "lon_decimal_calculada")

            n_puntos = 0
            _bounds = []
            for _, row in df_m.iterrows():
                lat = row.get(col_lat)
                lon = row.get(col_lon)
                if pd.isna(lat) or pd.isna(lon): continue
                try:
                    lat, lon = float(lat), float(lon)
                    if not (-5 <= lat <= 16 and -82 <= lon <= -60): continue
                except: continue

                val      = str(row.get(col_val, "")).strip() if col_val else ""
                color    = _val_color(val)
                vdesc    = _val_desc(val)

                # ── Datos del popup ──
                catalogo  = str(row.get("*Número de catálogo", row.get("Número de catálogo", row.get("catalogNumber", "")))).strip()
                catalogo  = catalogo if catalogo and catalogo.lower() != "nan" else "—"
                # Nombre científico: si existe la columna úsala; si no, construir Género + Epíteto
                _sci = str(row.get("Nombre científico", row.get("scientificName", ""))).strip()
                _gen = str(row.get("Género", row.get("genus", ""))).strip()
                _epi = str(row.get("Epíteto", row.get("specificEpithet", ""))).strip()
                if _sci and _sci.lower() != "nan":
                    especie = _sci
                else:
                    especie = (f"{_gen} {_epi}".strip()) or "—"
                especie = especie if especie.lower() != "nan" else "—"
                municipio = str(row.get("*Municipio",                row.get("county",         "—")))
                depto     = str(row.get("*Departamento",             row.get("stateProvince",  "—")))
                localidad = str(row.get("*Localidad estandarizada",  row.get("locality",       "—")))
                nivel_ini = row.get("Nivel_inicial", row.get("Nivel de calidad inicial", "—"))
                nivel_fin = row.get("Nivel_final",   row.get("Nivel de calidad final",   "—"))
                muni_det  = str(row.get("municipio_detectado", "—"))
                incert    = str(row.get("Incertidumbre de coordenadas (m)", row.get("coordinateUncertaintyInMeters", "—")))
                elev      = str(row.get("Elevación mínima (msnm)", row.get("minimumElevationInMeters", "—")))

                if "OK" in val or "✅" in val:
                    val_sym, val_bg, val_fg = "[✓]", "#E8F5E0", "#2D5016"
                elif "Revisar" in val or "⚠" in val:
                    val_sym, val_bg, val_fg = "[!]", "#FEF3C7", "#7C5A00"
                elif "Error" in val or "❌" in val:
                    val_sym, val_bg, val_fg = "[X]", "#FEE2E2", "#7F1D1D"
                else:
                    val_sym, val_bg, val_fg = "[~]", "#F3F4F6", "#6B7280"

                popup_html = f"""
                <div style="font-family:'DM Sans',sans-serif;min-width:240px;max-width:320px;font-size:13px;padding:6px 4px">
                  <div style="font-weight:700;font-size:14px;color:#1a1a1a;margin-bottom:2px">{catalogo}</div>
                  <div style="font-style:italic;color:#555;margin-bottom:10px;font-size:13px">{especie}</div>

                  <div style="background:#f0f4ec;border-radius:6px;padding:8px 10px;margin-bottom:8px">
                    <div style="font-weight:600;font-size:11px;color:#4A7C2F;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Localidad</div>
                    <div style="color:#333;font-size:12px">{localidad}</div>
                    <div style="margin-top:4px;color:#555;font-size:12px"><b>Municipio reportado:</b> {municipio}, {depto}</div>
                    <div style="color:#555;font-size:12px"><b>Municipio detectado GADM:</b> {muni_det}</div>
                  </div>

                  <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:8px">
                    <div style="background:#fafafa;border:1px solid #e8e8e8;border-radius:4px;padding:6px 8px">
                      <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:.05em">Nivel inicial</div>
                      <div style="font-size:16px;font-weight:600;color:#2D5016">{nivel_ini}</div>
                    </div>
                    <div style="background:#fafafa;border:1px solid #e8e8e8;border-radius:4px;padding:6px 8px">
                      <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:.05em">Nivel final</div>
                      <div style="font-size:16px;font-weight:600;color:#2D5016">{nivel_fin}</div>
                    </div>
                  </div>

                  <div style="background:#fafafa;border:1px solid #e8e8e8;border-radius:4px;padding:6px 10px;margin-bottom:8px;font-size:12px">
                    <div><b>Coordenadas finales:</b> {lat:.5f}, {lon:.5f}</div>
                    <div><b>Incertidumbre:</b> {incert} m</div>
                    <div><b>Elevación reportada:</b> {elev} msnm</div>
                  </div>

                  <div style="padding:6px 10px;border-radius:4px;font-size:12px;font-weight:600;
                              background:{val_bg};color:{val_fg}">
                    {val_sym} {vdesc}
                  </div>
                </div>"""

                if "OK" in val or "✅" in val:        _ck = "OK"
                elif "Revisar" in val or "⚠" in val:  _ck = "Revisar"
                elif "Error" in val or "❌" in val:    _ck = "Error"
                else:                                  _ck = "Sin"

                folium.CircleMarker(
                    location=[lat, lon], radius=6,
                    color="white", weight=1.5,
                    fill=True, fill_color=color, fill_opacity=0.9,
                    popup=folium.Popup(popup_html, max_width=320),
                    tooltip=f"{catalogo} · {especie} · Nivel {nivel_fin}",
                ).add_to(clusters[_ck])
                _bounds.append([lat, lon])
                n_puntos += 1

            folium.LayerControl(collapsed=False).add_to(m)
            if _bounds:
                m.fit_bounds(_bounds, padding=(30, 30))
            st.caption(f"Mostrando **{n_puntos}** puntos · 🟢 OK · 🟠 Revisar · 🔴 Error · ⚪ Sin validación. "
                       f"Los círculos de agrupación usan el mismo color que los puntos.")
            st_folium(m, width=None, height=520, returned_objects=[],
                      key="visor_mapa")

        except ImportError:
            st.warning("Instala streamlit-folium para activar el visor.")

    # ── Pestaña 2: Tabla ──────────────────────────
    with tab2:
        cols_ver = [
            "*Municipio","*Departamento","*Localidad estandarizada",
            "Nivel de calidad inicial","Nivel de calidad final",
            "Resultado validación espacial",
            "Latitud georreferenciada","Longitud georreferenciada",
            "Comentarios de la georreferenciación","Origen",
        ]
        cols_ok = [c for c in cols_ver if c in df.columns]

        niveles_disp = sorted(df["Nivel_final"].dropna().unique().tolist())
        filtro_niv   = st.multiselect(
            "Filtrar por nivel",
            [f"Nivel {int(n)}" for n in niveles_disp],
            default=[f"Nivel {int(n)}" for n in niveles_disp],
        )
        niveles_sel = [int(x.split()[-1]) for x in filtro_niv]
        df_t = df[df["Nivel_final"].isin(niveles_sel)][cols_ok] if cols_ok else df[df["Nivel_final"].isin(niveles_sel)]

        st.dataframe(df_t, use_container_width=True, height=500, hide_index=True)
        st.caption(f"{len(df_t)} de {total} registros")

    # ── Pestaña 3: Descarga ───────────────────────
    with tab3:
        st.markdown("""
        <div class="dl-card">
          <div class="dl-title">Reporte de georreferenciación</div>
          <div class="dl-desc">
            Archivo Excel con dos hojas: <b>Resumen</b> (estadísticas del proceso)
            y <b>Registros</b> (264 registros con colores, niveles y comentarios
            según el protocolo SiB Colombia).
          </div>
          <ul class="dl-list">
            <li>Coordenadas corregidas y validadas (Tabla 2 del manual)</li>
            <li>Nivel de calidad inicial y final (Tabla 9 del manual)</li>
            <li>Resultado de validación espacial por municipio (sección 3.6.1)</li>
            <li>Comentarios de georreferenciación por registro (Tabla 13)</li>
            <li>Campos obligatorios Darwin Core (Tabla 6 y 12)</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label="⬇️  Descargar reporte .xlsx",
            data=st.session_state.excel_bytes,
            file_name="georeferenciacion_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
            type="primary",
        )

        st.markdown("""
        <div class="section-head" style="margin-top:32px">
          <div class="section-label">Referencia metodológica</div>
          <div class="section-line"></div>
        </div>
        <div class="reference">
          Escobar D., Jojoa L.M., Díaz S.R., Rudas E., Albarracín R.D.,
          Ramírez C., Gómez J.Y., López C.R., Saavedra J., Ortiz R. (2016).
          Georreferenciación de localidades: Una guía de referencia para
          colecciones biológicas. Instituto de Investigación de Recursos
          Biológicos Alexander von Humboldt – Instituto de Ciencias Naturales,
          Universidad Nacional de Colombia. Bogotá D.C., Colombia. 144 p.
        </div>
        """, unsafe_allow_html=True)
