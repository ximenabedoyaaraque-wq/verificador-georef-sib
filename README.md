# Verificador de Georreferenciación — SiB Colombia

Herramienta para la validación y georreferenciación de localidades 
en colecciones biológicas, basada en el protocolo del Instituto de 
Investigación de Recursos Biológicos Alexander von Humboldt e 
Instituto de Ciencias Naturales (Universidad Nacional de Colombia).

## Autoría

**Ximena Bedoya Araque**  
Estudiante de Ecología — Universidad CES  
Pasantía en Colecciones Biológicas CBUCES  
Medellín, Colombia

© 2026 Ximena Bedoya Araque. Todos los derechos reservados.  
Uso libre para investigación y educación no comercial.  
Para uso comercial o redistribución, contactar al autor.

## Referencia metodológica

Escobar D., Jojoa L.M., Díaz S.R., Rudas E., Albarracín R.D., 
Ramírez C., Gómez J.Y., López C.R., Saavedra J., Ortiz R. (2016). 
*Georreferenciación de localidades: Una guía de referencia para 
colecciones biológicas.* Instituto Humboldt – ICN/UNAL. 
Bogotá D.C., Colombia. 144 p.

## Fuentes de datos geográficos

- Límites municipales y departamentales: GADM v4.1 (uso académico)
- Capas de visualización: John-Guerra (github.com/john-guerra), 
  basado en shapefiles del DANE
- Elevación: Open-Elevation API / Open-Topo-Data API

## Cómo usar

Accede directamente desde el navegador — no requiere instalación:

🔗 [Abrir la app](https://verificador-georef-sib.streamlit.app)

1. Sube tu base con coordenadas (.xlsx)
2. Sube tu base sin coordenadas (.xlsx)
3. Haz clic en "Ejecutar análisis"
4. Revisa los resultados en el visor de puntos
5. Descarga el reporte final en Excel

## Para desarrolladores

Si quieres correr el código localmente:

1. `pip install -r requirements.txt`
2. `streamlit run app.py`

## Estructura del proyecto

```
verificador-georef-sib/
├── app.py                  
├── bloques/                
│   ├── bloque1_coordenadas.py
│   ├── bloque2_municipio.py
│   ├── bloque3_elevacion.py
│   ├── bloque4_incertidumbre.py
│   ├── bloque5_campos.py
│   ├── bloque6_verbatim.py
│   ├── bloque7_clasificacion.py
│   ├── bloque8_reclasificacion.py
│   ├── bloque9_centroides.py
│   └── bloque10_exportar.py
├── datos/
│   └── gadm41_COL_2.json   
└── requirements.txt
```
