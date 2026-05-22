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
                aplicar_post_procesamiento,
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

            prog.progress(60, text="Verificando elevación…")
            df = aplicar_bloque3(df)

            prog.progress(70, text="Asignando centroides…")
            if gadm_ok and GADM_PATH:
                df = aplicar_bloque9(df, GADM_PATH, usar_nominatim=True)

            prog.progress(82, text="Post-procesamiento: incertidumbre, coordenadas y comentarios…")
            df = aplicar_post_procesamiento(df)

            # Sacar elevación API de columnas internas → columnas visibles en Excel
            if "elevacion_api" in df.columns:
                df["Elevación API (msnm)"]        = df["elevacion_api"]
                df["Validación elevación"]         = df["elevacion_estado"].map(
                    {"OK": "[✓] OK", "Revisar": "[!] Revisar"}).fillna("")
                df["Nota elevación"]               = df["elevacion_nota"]

            prog.progress(90, text="Generando reporte Excel…")
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

            c1, c2 = st.columns([2,2])
            with c1:
                filtro_res = st.selectbox("Resultado", ["Todos","✅ OK","⚠ Revisar","❌ Error"])
            with c2:
                deptos = ["Todos"] + sorted(df["*Departamento"].dropna().unique().tolist()) \
                    if "*Departamento" in df.columns else ["Todos"]
                filtro_dep = st.selectbox("Departamento", deptos)

            df_m = df.copy()
            # filtro por resultado: usar contains para soportar ambos formatos
            if filtro_res != "Todos" and col_val:
                kw = {"✅ OK":"OK","⚠ Revisar":"Revisar","❌ Error":"Error"}.get(filtro_res,"")
                if kw:
                    df_m = df_m[df_m[col_val].str.contains(kw, na=False)]
            if filtro_dep != "Todos" and "*Departamento" in df_m.columns:
                df_m = df_m[df_m["*Departamento"] == filtro_dep]

            m = folium.Map(location=[5.5,-74.5], zoom_start=6,
                           tiles="CartoDB Positron", prefer_canvas=True)
            folium.TileLayer(
                "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
                attr="Esri", name="Satelital", overlay=False, control=True
            ).add_to(m)

            cluster = MarkerCluster(
                options={"maxClusterRadius":40,"disableClusteringAtZoom":12}
            ).add_to(m)

            col_lat = next((c for c in ["Latitud georreferenciada", "lat_wgs84", "lat_decimal_calculada"] if c in df_m.columns), "lat_decimal_calculada")
            col_lon = next((c for c in ["Longitud georreferenciada", "lon_wgs84", "lon_decimal_calculada"] if c in df_m.columns), "lon_decimal_calculada")

            n_puntos = 0
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

                # ── Datos del popup (con fallbacks robustos) ──
                catalogo  = str(row.get("Número de catálogo",    row.get("catalogNumber",    "—")))
                especie   = str(row.get("Nombre científico",     row.get("scientificName",   "—")))
                municipio = str(row.get("*Municipio",            row.get("county",           "—")))
                depto     = str(row.get("*Departamento",         row.get("stateProvince",    "—")))
                localidad = str(row.get("*Localidad estandarizada", row.get("locality",     "—")))
                nivel_ini = row.get("Nivel_inicial", row.get("Nivel de calidad inicial", "—"))
                nivel_fin = row.get("Nivel_final",   row.get("Nivel de calidad final",   "—"))
                muni_det  = str(row.get("municipio_detectado", "—"))
                depto_det = str(row.get("depto_detectado",     "—"))
                incert    = str(row.get("Incertidumbre de coordenadas (m)", row.get("coordinateUncertaintyInMeters", "—")))
                elev      = str(row.get("Elevación mínima (msnm)", row.get("minimumElevationInMeters", "—")))
                elev_api  = str(row.get("elevacion_api", "—"))
                elev_est  = str(row.get("elevacion_estado", ""))
                elev_nota = str(row.get("elevacion_nota", ""))
                fecha     = str(row.get("Fecha",     row.get("eventDate", "—")))
                colector  = str(row.get("Colector",  row.get("recordedBy", "—")))
                origen    = str(row.get("Origen",    "—"))

                popup_html = f"""
                <div style="font-family:'DM Sans',sans-serif;min-width:240px;max-width:300px;font-size:13px;padding:6px 4px">
                  <div style="font-weight:700;font-size:14px;color:#1a1a1a;margin-bottom:2px">{catalogo}</div>
                  <div style="font-style:italic;color:#555;margin-bottom:10px;font-size:13px">{especie}</div>

                  <div style="background:#f0f4ec;border-radius:6px;padding:8px 10px;margin-bottom:8px">
                    <div style="font-weight:600;font-size:11px;color:#4A7C2F;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Localidad</div>
                    <div style="color:#333;font-size:12px">{localidad}</div>
                    <div style="margin-top:4px;color:#666;font-size:12px">{municipio}, {depto}</div>
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
                    <div><b>Coordenadas:</b> {lat:.5f}, {lon:.5f}</div>
                    <div><b>Incertidumbre:</b> {incert} m</div>
                    <div><b>Municipio detectado:</b> {muni_det}, {depto_det}</div>
                    <div><b>Elevación reportada:</b> {elev} msnm</div>
                    <div><b>Elevación API:</b> {elev_api} msnm {("⚠" if elev_est=="Revisar" else "✓") if elev_api != "—" else ""}</div>
                  </div>

                  <div style="padding:5px 10px;border-radius:4px;font-size:11px;font-weight:600;
                              background:{'#E8F5E0' if 'OK' in vdesc else '#FEF3C7' if 'vecino' in vdesc else '#FEE2E2'};
                              color:{'#2D5016' if 'OK' in vdesc else '#7C5A00' if 'vecino' in vdesc else '#7F1D1D'}">
                    {val} — {vdesc}
                  </div>

                  <div style="margin-top:8px;font-size:11px;color:#aaa">
                    {fecha} · {colector} · {origen}
                  </div>
                </div>"""

                folium.CircleMarker(
                    location=[lat, lon], radius=6,
                    color="white", weight=1.5,
                    fill=True, fill_color=color, fill_opacity=0.9,
                    popup=folium.Popup(popup_html, max_width=320),
                    tooltip=f"{catalogo} · {especie} · Nivel {nivel_fin}",
                ).add_to(cluster)
                n_puntos += 1

            folium.LayerControl().add_to(m)
            st.caption(f"Mostrando **{n_puntos}** puntos · Verde = OK · Naranja = Revisar · Rojo = Error")
            st_folium(m, width="100%", height=520, returned_objects=[])

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
