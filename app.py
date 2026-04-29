"""
Verificador de Georreferenciación — SiB Colombia
© 2026 Ximena Bedoya Araque · Universidad CES · Colecciones Biológicas CBUCES
Protocolo: Escobar D. et al. (2016) · Instituto Humboldt / ICN-UNAL
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
import sys
import io

# ─── Configuración de página ──────────────────────────────────
st.set_page_config(
    page_title="Verificador de Georreferenciación · SiB Colombia",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS: modo oscuro/claro automático + diseño limpio ────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

/* Variables de color — se adaptan al tema del sistema */
:root {
    --c-primary:   #1B4332;
    --c-accent:    #40916C;
    --c-light:     #D8F3DC;
    --c-warn:      #F4A261;
    --c-error:     #E63946;
    --c-text:      #1A1A2E;
    --c-muted:     #6B7280;
    --c-surface:   #F8FAF9;
    --c-border:    #E2E8E4;
    --radius:      10px;
    --font:        'DM Sans', sans-serif;
    --font-mono:   'DM Mono', monospace;
}

@media (prefers-color-scheme: dark) {
    :root {
        --c-primary:   #74C69D;
        --c-accent:    #52B788;
        --c-light:     #1B4332;
        --c-warn:      #F4A261;
        --c-error:     #FF6B6B;
        --c-text:      #E8F5E9;
        --c-muted:     #9CA3AF;
        --c-surface:   #1A1F1C;
        --c-border:    #2D3B35;
    }
}

html, body, [class*="css"] {
    font-family: var(--font) !important;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: var(--c-surface);
    border-right: 1px solid var(--c-border);
}

/* Cards de métricas */
.metric-card {
    background: var(--c-surface);
    border: 1px solid var(--c-border);
    border-radius: var(--radius);
    padding: 16px 18px;
    text-align: center;
}
.metric-num {
    font-size: 28px;
    font-weight: 600;
    color: var(--c-primary);
    line-height: 1.1;
}
.metric-label {
    font-size: 11px;
    color: var(--c-muted);
    margin-top: 3px;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

/* Badges de resultado */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 500;
}
.badge-ok      { background: #D8F3DC; color: #1B4332; }
.badge-warn    { background: #FFF3CD; color: #7C5A00; }
.badge-error   { background: #FFE5E5; color: #8B0000; }
.badge-nd      { background: #F1F1F1; color: #555; }

/* Título de sección */
.section-title {
    font-size: 13px;
    font-weight: 500;
    color: var(--c-muted);
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin: 24px 0 12px;
    padding-bottom: 6px;
    border-bottom: 1px solid var(--c-border);
}

/* Upload zone */
.upload-label {
    font-size: 12px;
    font-weight: 500;
    color: var(--c-muted);
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 4px;
}

/* Copyright */
.copyright {
    font-size: 10px;
    color: var(--c-muted);
    line-height: 1.6;
    margin-top: 24px;
    padding-top: 12px;
    border-top: 1px solid var(--c-border);
}

/* Progress step */
.step-row {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 6px;
    font-size: 13px;
}
.step-dot {
    width: 8px; height: 8px;
    border-radius: 50%;
    flex-shrink: 0;
}
.dot-ok   { background: #40916C; }
.dot-pend { background: #D1D5DB; }
.dot-run  { background: #F4A261; }
</style>
""", unsafe_allow_html=True)

# ─── Rutas ────────────────────────────────────────────────────
BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
GADM_PATH = os.path.join(BASE_DIR, "datos", "gadm41_COL_2.json")
sys.path.insert(0, BASE_DIR)
sys.path.insert(0, os.path.join(BASE_DIR, "bloques"))

# ─── Sidebar ──────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🌿 Verificador de Georreferenciación")
    st.markdown("**SiB Colombia · Instituto Humboldt**")
    st.divider()

    st.markdown('<div class="upload-label">Base con coordenadas (.xlsx)</div>',
                unsafe_allow_html=True)
    file_180 = st.file_uploader("", type=["xlsx","xls"],
                                 key="file_180", label_visibility="collapsed")

    st.markdown('<div class="upload-label" style="margin-top:12px">Base sin coordenadas (.xlsx)</div>',
                unsafe_allow_html=True)
    file_84 = st.file_uploader("", type=["xlsx","xls"],
                                key="file_84", label_visibility="collapsed")

    st.divider()

    gadm_ok = os.path.exists(GADM_PATH)
    if gadm_ok:
        st.success("✓ Capa GADM Colombia cargada", icon="🗺️")
    else:
        st.warning("Capa GADM no encontrada en datos/", icon="⚠️")
        st.caption("Sube gadm41_COL_2.json a la carpeta datos/ del repositorio.")

    ejecutar = st.button("▶ Ejecutar análisis",
                         use_container_width=True,
                         type="primary",
                         disabled=(file_180 is None or file_84 is None))

    st.markdown("""
    <div class="copyright">
        © 2026 Ximena Bedoya Araque<br>
        Estudiante de Ecología · Universidad CES<br>
        Pasantía en Colecciones Biológicas CBUCES<br>
        Medellín, Colombia<br><br>
        Basado en: Escobar D. et al. (2016)<br>
        Instituto Humboldt – ICN/UNAL<br>
        Licencia CC BY-NC 4.0
    </div>
    """, unsafe_allow_html=True)

# ─── Estado de sesión ─────────────────────────────────────────
if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = None
if "procesado" not in st.session_state:
    st.session_state.procesado = False

# ─── Procesar cuando se presiona el botón ─────────────────────
if ejecutar and file_180 is not None and file_84 is not None:
    with st.spinner("Procesando registros..."):
        try:
            from verificador_georef_completo_1 import (
                aplicar_bloque1, aplicar_bloque5, aplicar_bloque6,
                aplicar_bloque7, aplicar_bloque8, aplicar_bloque9,
                aplicar_bloque10
            )

            # Guardar archivos temporalmente
            with open("/tmp/base_180.xlsx", "wb") as f:
                f.write(file_180.read())
            with open("/tmp/base_84.xlsx", "wb") as f:
                f.write(file_84.read())

            # Bloque 1 — leer, unir, limpiar coordenadas
            progress = st.progress(0, text="Estandarizando coordenadas...")
            df = aplicar_bloque1("/tmp/base_84.xlsx", "/tmp/base_180.xlsx")

            # Bloque 5 — campos obligatorios
            progress.progress(20, text="Verificando campos DwC...")
            df, _ = aplicar_bloque5(df)

            # Bloque 7 — clasificación niveles
            progress.progress(35, text="Clasificando niveles de calidad...")
            df = aplicar_bloque7(df)

            # Bloque 8 — validación espacial (requiere GADM)
            progress.progress(50, text="Validando coordenadas contra municipios...")
            if gadm_ok:
                df = aplicar_bloque8(df, GADM_PATH)
            else:
                df["Nivel_final"]         = df["Nivel_inicial"].copy()
                df["lat_wgs84"]           = df["lat_decimal_calculada"]
                df["lon_wgs84"]           = df["lon_decimal_calculada"]
                df["validacion_b2"]       = df["conversion_estado"].map(
                    {"OK":"OK","Revisar":"Revisar","Error":"Error","sin coordenadas":""}
                ).fillna("")
                df["municipio_detectado"] = ""
                df["depto_detectado"]     = ""
                df["mensaje_b2"]          = ""

            # Bloque 9 — centroides para sin coordenadas
            progress.progress(70, text="Asignando centroides...")
            if gadm_ok:
                df = aplicar_bloque9(df, GADM_PATH, usar_nominatim=True)

            progress.progress(90, text="Generando reporte...")
            st.session_state.df_resultado = df
            st.session_state.excel_bytes  = aplicar_bloque10(df, idioma=None)
            st.session_state.procesado    = True
            progress.progress(100, text="¡Listo!")

        except Exception as e:
            st.error(f"Error durante el procesamiento: {e}")
            import traceback
            st.code(traceback.format_exc())

# ─── Pantalla de bienvenida ───────────────────────────────────
if not st.session_state.procesado:
    st.markdown("# Verificador de Georreferenciación")
    st.markdown(
        "Herramienta para la validación y georreferenciación de localidades "
        "en colecciones biológicas, basada en el protocolo del **Instituto Humboldt** "
        "e **Instituto de Ciencias Naturales (UNAL)**."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-num">10</div>
            <div class="metric-label">Procesos automatizados</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-num">7</div>
            <div class="metric-label">Niveles de calidad</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-num">50k+</div>
            <div class="metric-label">Registros soportados</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="section-title">Cómo usar</div>', unsafe_allow_html=True)
    pasos = [
        ("Sube tu base con coordenadas (.xlsx)", True),
        ("Sube tu base sin coordenadas (.xlsx)", True),
        ("Presiona ▶ Ejecutar análisis", False),
        ("Revisa los resultados en el visor de puntos", False),
        ("Descarga el reporte en Excel", False),
    ]
    for texto, _ in pasos:
        st.markdown(f"→ {texto}")

    st.info(
        "💡 Ambas bases deben venir estandarizadas del proceso previo "
        "(Taller 1 — Verificador de Localidades), con la columna "
        "**\\*Localidad estandarizada** ya corregida.",
        icon="ℹ️"
    )

# ─── Resultados ───────────────────────────────────────────────
else:
    df = st.session_state.df_resultado

    # Métricas resumen
    total     = len(df)
    n1        = int((df["Nivel_final"] == 1).sum())
    n2_6      = int(df["Nivel_final"].isin([2,3,4,5,6]).sum())
    n7        = int((df["Nivel_final"] == 7).sum())

    val_ok    = int((df.get("validacion_b2","") == "OK").sum()) if "validacion_b2" in df.columns else 0
    val_rev   = int((df.get("validacion_b2","") == "Revisar").sum()) if "validacion_b2" in df.columns else 0
    val_err   = int((df.get("validacion_b2","") == "Error").sum()) if "validacion_b2" in df.columns else 0

    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    metricas = [
        (c1, total,  "Total registros"),
        (c2, n1,     "Con coordenadas"),
        (c3, n2_6,   "Georreferenciados"),
        (c4, n7,     "Nivel 7"),
        (c5, val_ok, "[✓] OK"),
        (c6, val_rev,"[!] Revisar"),
        (c7, val_err,"[X] Error"),
    ]
    for col, num, label in metricas:
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-num">{num}</div>
                <div class="metric-label">{label}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("")

    # Pestañas
    tab1, tab2, tab3 = st.tabs(["🗺️ Visor de puntos", "📋 Tabla de resultados", "⬇️ Descargar reporte"])

    # ── Pestaña 1: Visor de puntos ────────────────────────────
    with tab1:
        try:
            import folium
            from streamlit_folium import st_folium
            from folium.plugins import MarkerCluster

            # Filtros
            col_f1, col_f2 = st.columns([2, 2])
            with col_f1:
                filtro_resultado = st.selectbox(
                    "Filtrar por resultado",
                    ["Todos", "[✓] OK", "[!] Revisar", "[X] Error", "Sin validación"]
                )
            with col_f2:
                deptos = ["Todos"] + sorted(df["*Departamento"].dropna().unique().tolist()) \
                    if "*Departamento" in df.columns else ["Todos"]
                filtro_depto = st.selectbox("Filtrar por departamento", deptos)

            # Aplicar filtros
            df_mapa = df.copy()
            if filtro_resultado != "Todos" and "validacion_b2" in df_mapa.columns:
                mapa_val = {"[✓] OK":"OK","[!] Revisar":"Revisar","[X] Error":"Error","Sin validación":""}
                df_mapa = df_mapa[df_mapa["validacion_b2"] == mapa_val.get(filtro_resultado,"")]
            if filtro_depto != "Todos" and "*Departamento" in df_mapa.columns:
                df_mapa = df_mapa[df_mapa["*Departamento"] == filtro_depto]

            # Construir mapa
            m = folium.Map(
                location=[5.5, -74.5],
                zoom_start=6,
                tiles="OpenStreetMap",
                prefer_canvas=True,
            )
            # Capa satelital opcional
            folium.TileLayer(
                tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
                attr="Esri",
                name="Satelital",
                overlay=False,
                control=True,
            ).add_to(m)

            # Usar MarkerCluster para muchos puntos
            cluster = MarkerCluster(
                options={"maxClusterRadius": 40, "disableClusteringAtZoom": 12}
            ).add_to(m)

            col_lat = "lat_wgs84" if "lat_wgs84" in df_mapa.columns else "lat_decimal_calculada"
            col_lon = "lon_wgs84" if "lon_wgs84" in df_mapa.columns else "lon_decimal_calculada"
            col_val = "validacion_b2"

            colores_val = {"OK":"green","Revisar":"orange","Error":"red","":"gray"}

            puntos_agregados = 0
            for _, row in df_mapa.iterrows():
                lat = row.get(col_lat)
                lon = row.get(col_lon)
                if pd.isna(lat) or pd.isna(lon): continue
                try:
                    lat, lon = float(lat), float(lon)
                    if not (-5 <= lat <= 16 and -82 <= lon <= -60): continue
                except: continue

                val      = str(row.get(col_val,"")).strip()
                color    = colores_val.get(val, "gray")
                especie  = str(row.get("Nombre científico", row.get("scientificName","—")))
                municipio= str(row.get("*Municipio", row.get("county","—")))
                depto    = str(row.get("*Departamento", row.get("stateProvince","—")))
                nivel    = row.get("Nivel_final", "—")
                catalogo = str(row.get("Número de catálogo", row.get("catalogNumber","—")))
                muni_det = str(row.get("municipio_detectado","—"))
                val_txt  = {"OK":"✓ Dentro del municipio","Revisar":"⚠ Municipio vecino","Error":"✗ Fuera de Colombia"}.get(val,"— Sin validación")

                popup_html = f"""
                <div style="font-family:DM Sans,sans-serif;min-width:220px;font-size:13px">
                    <b style="font-size:14px">{catalogo}</b><br>
                    <i style="color:#555">{especie}</i>
                    <hr style="margin:6px 0;border:none;border-top:1px solid #eee">
                    <b>Municipio reportado:</b> {municipio}, {depto}<br>
                    <b>Municipio detectado:</b> {muni_det}<br>
                    <b>Lat:</b> {lat:.6f} · <b>Lon:</b> {lon:.6f}<br>
                    <b>Nivel:</b> {nivel}<br>
                    <b>Validación:</b> {val_txt}
                </div>
                """
                folium.CircleMarker(
                    location=[lat, lon],
                    radius=6,
                    color="white",
                    weight=1.5,
                    fill=True,
                    fill_color=color,
                    fill_opacity=0.85,
                    popup=folium.Popup(popup_html, max_width=280),
                    tooltip=f"{catalogo} — {especie}",
                ).add_to(cluster)
                puntos_agregados += 1

            folium.LayerControl().add_to(m)

            st.caption(f"Mostrando {puntos_agregados} puntos con coordenadas")
            st_folium(m, width="100%", height=520, returned_objects=[])

        except ImportError:
            st.warning("Instala streamlit-folium para ver el mapa: pip install streamlit-folium")

    # ── Pestaña 2: Tabla de resultados ────────────────────────
    with tab2:
        col_res = [
            "*Municipio","*Departamento","*Localidad estandarizada",
            "Nivel de calidad inicial","Nivel de calidad final",
            "Resultado validación espacial",
            "Latitud georreferenciada","Longitud georreferenciada",
            "Comentarios de la georreferenciación",
            "Origen"
        ]
        cols_presentes = [c for c in col_res if c in df.columns]

        # Filtro de nivel
        niveles_disponibles = sorted(df["Nivel_final"].dropna().unique().tolist())
        filtro_nivel = st.multiselect(
            "Filtrar por nivel de calidad",
            options=[f"Nivel {int(n)}" for n in niveles_disponibles],
            default=[f"Nivel {int(n)}" for n in niveles_disponibles],
        )
        niveles_sel = [int(x.split()[-1]) for x in filtro_nivel]
        df_tabla = df[df["Nivel_final"].isin(niveles_sel)][cols_presentes]

        st.dataframe(
            df_tabla,
            use_container_width=True,
            height=480,
            hide_index=True,
        )
        st.caption(f"{len(df_tabla)} registros mostrados de {total} totales")

    # ── Pestaña 3: Descarga ───────────────────────────────────
    with tab3:
        st.markdown("### Descarga el reporte completo")
        st.markdown(
            "El archivo Excel contiene dos hojas: "
            "**Resumen** con estadísticas del proceso y "
            "**Registros** con los 264 registros procesados, "
            "colores por resultado y comentarios de georreferenciación."
        )

        col_d1, col_d2 = st.columns([1,2])
        with col_d1:
            st.download_button(
                label="⬇️ Descargar reporte .xlsx",
                data=st.session_state.excel_bytes,
                file_name="georeferenciacion_resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

        with col_d2:
            st.markdown("""
            **El reporte incluye:**
            - Coordenadas corregidas y validadas
            - Nivel de calidad inicial y final (Tabla 9 del manual)
            - Resultado de validación espacial por municipio
            - Comentarios de georreferenciación (sección 3.6, manual)
            - Campos obligatorios DwC (Tabla 6 y 12 del manual)
            """)

        st.divider()
        st.markdown("**Referencia metodológica**")
        st.caption(
            "Escobar D., Jojoa L.M., Díaz S.R., Rudas E., Albarracín R.D., "
            "Ramírez C., Gómez J.Y., López C.R., Saavedra J., Ortiz R. (2016). "
            "Georreferenciación de localidades: Una guía de referencia para "
            "colecciones biológicas. Instituto Humboldt – ICN/UNAL. "
            "Bogotá D.C., Colombia. 144 p."
        )

