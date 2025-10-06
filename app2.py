import streamlit as st
import pandas as pd
from pptx import Presentation
from geopy.distance import geodesic
from datetime import datetime
import folium
from folium.plugins import MarkerCluster
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from copy import deepcopy
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
import re
import io

# ================================
# Funciones de lógica de negocio
# ================================

def generar_folio():
    """Genera un folio único con formato PROP-YYYYMMDD-XXX"""
    hoy = datetime.now().strftime("%Y%m%d")
    
    # Inicializar contador en session_state si no existe
    if 'contador_descargas' not in st.session_state:
        st.session_state.contador_descargas = 1
    
    # Obtener el número actual y incrementar
    numero = st.session_state.contador_descargas
    folio = f"PROP-{hoy}-{numero:03d}"
    return folio

def incrementar_folio():
    """Incrementa el contador de folios solo cuando se realiza una descarga"""
    if 'contador_descargas' not in st.session_state:
        st.session_state.contador_descargas = 1
    else:
        st.session_state.contador_descargas += 1

def analizar_formato_coordenada(coord_str):
    if pd.isna(coord_str) or coord_str == "" or str(coord_str).strip() == "":
        return "vacía", None, None, None
    coord_str = str(coord_str).strip()
    patron_grados_dir = r'(\d+)°\s*(\d+)\'\s*(\d+\.?\d*)\"\s*([NSWE])'
    coincidencia = re.search(patron_grados_dir, coord_str, re.IGNORECASE)
    if coincidencia:
        grados = float(coincidencia.group(1))
        minutos = float(coincidencia.group(2))
        segundos = float(coincidencia.group(3))
        direccion = coincidencia.group(4).upper()
        decimal = grados + minutos/60 + segundos/3600
        if direccion in ['S', 'W']: decimal = -decimal
        return "grados_dir", decimal, direccion, None
    patron_grados_sig = r'([+-]?\d+\.\d+)'
    coincidencia = re.search(patron_grados_sig, coord_str)
    if coincidencia:
        decimal = float(coincidencia.group(1))
        return "grados_sig", decimal, None, None
    patron_dms = r'([+-]?\d+)\s+(\d+)\s+(\d+\.?\d*)'
    coincidencia = re.search(patron_dms, coord_str)
    if coincidencia:
        grados = float(coincidencia.group(1))
        minutos = float(coincidencia.group(2))
        segundos = float(coincidencia.group(3))
        decimal = abs(grados) + minutos/60 + segundos/3600
        if grados < 0: decimal = -decimal
        return "dms", decimal, None, None
    patron_dm = r'([+-]?\d+)°\s*(\d+\.\d+)'
    coincidencia = re.search(patron_dm, coord_str)
    if coincidencia:
        grados = float(coincidencia.group(1))
        minutos = float(coincidencia.group(2))
        decimal = abs(grados) + minutos/60
        if grados < 0: decimal = -decimal
        return "dm", decimal, None, None
    if "," in coord_str and "." not in coord_str:
        try:
            coord_europeo = coord_str.replace(",", ".")
            decimal = float(coord_europeo)
            return "decimal_eu", decimal, None, None
        except ValueError:
            pass
    if coord_str.count(",") >= 2:
        try:
            coord_sin_comas = coord_str.replace(",", "")
            decimal = float(coord_sin_comas)
            return "decimal_miles", decimal, None, None
        except ValueError:
            pass
    try:
        decimal = float(coord_str)
        return "decimal", decimal, None, None
    except ValueError:
        pass
    return "desconocido", None, None, None

def ajustar_valor_utm(valor, es_latitud=True):
    if valor is None: return None
    valor_abs = abs(valor)
    if es_latitud and (-90 <= valor <= 90): return valor
    elif not es_latitud and (-180 <= valor <= 180): return valor
    if valor_abs > 180:
        if valor_abs >= 1000000000: divisor = 10000000
        elif valor_abs >= 100000000: divisor = 1000000
        elif valor_abs >= 10000000: divisor = 100000
        elif valor_abs >= 1000000: divisor = 10000
        elif valor_abs >= 100000: divisor = 1000
        else: divisor = 1000
        valor_ajustado = valor / divisor
        if es_latitud and (-90 <= valor_ajustado <= 90): return valor_ajustado
        elif not es_latitud and (-180 <= valor_ajustado <= 180): return valor_ajustado
        else: return valor / (divisor * 10)
    return valor

def estandarizar_coordenada_universal(coord_str, formato, valor_raw=None, direccion=None, es_latitud=True):
    if pd.isna(coord_str) or coord_str == "": return None
    coord_str = str(coord_str).strip()
    try:
        if formato in ["grados_dir", "grados_sig", "dms", "dm"] and valor_raw is not None:
            valor = valor_raw
        elif formato == "decimal_eu":
            valor = float(coord_str.replace(",", "."))
        elif formato == "decimal_miles":
            valor = float(coord_str.replace(",", ""))
        elif formato == "decimal":
            valor = float(coord_str)
        elif formato == "desconocido":
            numeros = re.findall(r'[+-]?\d+\.?\d*', coord_str)
            valor = float(numeros[0]) if numeros else None
        else:
            valor = None
        if valor is None: return None
        return ajustar_valor_utm(valor, es_latitud)
    except (ValueError, TypeError):
        return None

def detectar_inversion_universal(lat_str, lon_str):
    formato_lat, valor_lat, dir_lat, _ = analizar_formato_coordenada(lat_str)
    formato_lon, valor_lon, dir_lon, _ = analizar_formato_coordenada(lon_str)
    if valor_lat is None or valor_lon is None: return False, formato_lat, formato_lon, ["no_analizable"]
    valor_abs_lat, valor_abs_lon = abs(valor_lat), abs(valor_lon)
    es_patron_utm_invertido = ((valor_abs_lat > 1000000 or valor_abs_lon > 1000000) and
                              (valor_abs_lat < 1000000000 and valor_abs_lon < 1000000000) and
                              (valor_lat < 0 and valor_lon > 0))
    if es_patron_utm_invertido: return True, formato_lat, formato_lon, ["patron_utm_invertido"]
    rango_lat_absoluto, rango_lon_absoluto = (-90, 90), (-180, 180)
    lat_en_rango_valido = rango_lat_absoluto[0] <= valor_lat <= rango_lat_absoluto[1]
    lon_en_rango_valido = rango_lon_absoluto[0] <= valor_lon <= rango_lon_absoluto[1]
    direcciones_invertidas = (dir_lat in ['E', 'W'] and dir_lon in ['N', 'S']) if dir_lat and dir_lon else False
    lat_podria_ser_lon = (rango_lon_absoluto[0] <= valor_lat <= rango_lon_absoluto[1] and not lat_en_rango_valido)
    lon_podria_ser_lat = (rango_lat_absoluto[0] <= valor_lon <= rango_lat_absoluto[1] and not lon_en_rango_valido)
    diferencia_magnitud = abs(abs(valor_lat) - abs(valor_lon))
    lat_mayor_que_lon = abs(valor_lat) > abs(valor_lon) + 10
    signos_atipicos = (valor_lat < 0 and valor_lon > 0)
    criterios = []
    if direcciones_invertidas: criterios.append("direcciones_cardinales")
    if not lat_en_rango_valido and not lon_en_rango_valido and lat_podria_ser_lon and lon_podria_ser_lat: criterios.append("valores_fuera_de_rango")
    if lat_mayor_que_lon and diferencia_magnitud > 20: criterios.append("diferencia_magnitud")
    if signos_atipicos: criterios.append("signos_atipicos")
    probable_inversion = len(criterios) >= 2
    return probable_inversion, formato_lat, formato_lon, criterios

def duplicar_slide(prs, slide):
    new_slide = prs.slides.add_slide(slide.slide_layout)
    for shp in slide.shapes:
        el = deepcopy(shp.element)
        new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')
    for rel in slide.part.rels.values():
        try:
            if rel.reltype == RT.IMAGE:
                image_part = rel.target_part
                new_slide.part.relate_to(image_part, RT.IMAGE)
            elif rel.reltype == RT.HYPERLINK:
                new_slide.part.relate_to(rel.target_ref, RT.HYPERLINK)
        except Exception as e:
            st.error(f"⚠️ No se pudo copiar relación: {e}")
    return new_slide

def reemplazar_texto_slide(slide, fila, df_filtrado):
    """
    Reemplaza los marcadores de texto en una diapositiva con los valores de una fila.
    Crea un hipervínculo de Street View para el campo LATITUD.
    """
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                for campo in df_filtrado.columns:
                    marcador = f"{{{{{campo.strip()}}}}}"
                    
                    if marcador in run.text:
                        valor = fila.get(campo, "")
                        
                        if campo.strip().upper() == "LATITUD":
                            try:
                                run.text = run.text.replace(marcador, str(valor))
                                run.hyperlink.address = fila.get("STREET_VIEW", "#")
                            except Exception as e:
                                st.warning(f"⚠️ No se pudo crear el hipervínculo para {campo}: {e}")
                        else:
                            run.text = run.text.replace(marcador, str(valor))

# ================================
# ESTRUCTURA DE LA APP STREAMLIT
# ================================

st.set_page_config(layout="wide")
st.title("🤖 Analizador de Espectaculares")
st.subheader("¡Hola! Soy tu asistente virtual para encontrar los mejores espectaculares publicitarios para tu negocio.")
st.write("---")

# Inicializar st.session_state
if 'df_filtrado' not in st.session_state:
    st.session_state.df_filtrado = pd.DataFrame()
if 'uploaded_df' not in st.session_state:
    st.session_state.uploaded_df = None
if 'negocio_lat' not in st.session_state:
    st.session_state.negocio_lat = 19.4326
if 'negocio_lon' not in st.session_state:
    st.session_state.negocio_lon = -99.1332
if 'presupuesto_min' not in st.session_state:
    st.session_state.presupuesto_min = 0.0
if 'presupuesto_max' not in st.session_state:
    st.session_state.presupuesto_max = 100000.0
if 'radio_km' not in st.session_state:
    st.session_state.radio_km = 5.0
if 'contador_descargas' not in st.session_state:
    st.session_state.contador_descargas = 1
if 'folio_actual' not in st.session_state:
    st.session_state.folio_actual = generar_folio()
if 'busqueda_realizada' not in st.session_state:
    st.session_state.busqueda_realizada = False
    
# 1. UPLOAD CSV
uploaded_file = st.file_uploader("📂 **Paso 1: Sube tu archivo CSV de inventario**", type="csv")
if uploaded_file:
    st.session_state.uploaded_df = pd.read_csv(uploaded_file, sep=",")
    st.session_state.uploaded_df.columns = st.session_state.uploaded_df.columns.str.strip()
    st.success(f"✅ CSV cargado con **{len(st.session_state.uploaded_df)}** registros.")
    if "TARIFA PUBLICO" not in st.session_state.uploaded_df.columns:
        st.error("❌ No se encuentra la columna **'TARIFA PUBLICO'** en el CSV.")
        st.session_state.uploaded_df = None
        st.stop()
    else:
        st.session_state.uploaded_df["TARIFA"] = (st.session_state.uploaded_df["TARIFA PUBLICO"].astype(str).str.replace(r"[^\d.]", "", regex=True).replace("", "0").astype(float))


# 2. INPUTS PARA LA BÚSQUEDA
st.write("---")
st.header("🎯 **Paso 2: Define tus criterios de búsqueda**")
col1, col2, col3 = st.columns(3)

with col1:
    lat_input = st.text_input("🧭 Latitud del negocio:", value=str(st.session_state.negocio_lat), key='negocio_lat_text_input')
    try:
        st.session_state.negocio_lat = float(lat_input) if lat_input.strip() != "" else 19.4326
    except ValueError:
        st.warning("⚠️ Formato de latitud incorrecto. Usando valor predetermeterminado.")
        st.session_state.negocio_lat = 19.4326

    st.session_state.presupuesto_min = st.number_input("💰 Presupuesto mínimo:", value=st.session_state.presupuesto_min, format="%.2f", key='presupuesto_min_input')

with col2:
    lon_input = st.text_input("🧭 Longitud del negocio:", value=str(st.session_state.negocio_lon), key='negocio_lon_text_input')
    try:
        st.session_state.negocio_lon = float(lon_input) if lon_input.strip() != "" else -99.1332
    except ValueError:
        st.warning("⚠️ Formato de longitud incorrecto. Usando valor predeterminado.")
        st.session_state.negocio_lon = -99.1332
    
    st.session_state.presupuesto_max = st.number_input("💰 Presupuesto máximo:", value=st.session_state.presupuesto_max, format="%.2f", key='presupuesto_max_input')

with col3:
    st.session_state.radio_km = st.slider("📏 Radio de búsqueda (km):", min_value=0.5, max_value=50.0, value=st.session_state.radio_km, step=0.5, key='radio_km_input')

# 3. FILTRADO Y GENERACIÓN DE RESULTADOS
if st.button("🚀 **Iniciar Búsqueda**") and st.session_state.uploaded_df is not None:
    with st.spinner("🔄 Analizando y corrigiendo coordenadas..."):
        df_copy = st.session_state.uploaded_df.copy()
        coordenadas_corregidas = []
        for _, row in df_copy.iterrows():
            lat_original, lon_original = row["LATITUD"], row["LONGITUD"]
            formato_lat, valor_lat, dir_lat, _ = analizar_formato_coordenada(lat_original)
            formato_lon, valor_lon, dir_lon, _ = analizar_formato_coordenada(lon_original)
            invertidas, f_lat_det, f_lon_det, criterios = detectar_inversion_universal(lat_original, lon_original)
            if invertidas:
                lat_corregida = estandarizar_coordenada_universal(lon_original, f_lon_det, valor_lon, dir_lon, True)
                lon_corregida = estandarizar_coordenada_universal(lat_original, f_lat_det, valor_lat, dir_lat, False)
            else:
                lat_corregida = estandarizar_coordenada_universal(lat_original, formato_lat, valor_lat, dir_lat, True)
                lon_corregida = estandarizar_coordenada_universal(lon_original, formato_lon, valor_lon, dir_lon, False)
            coordenadas_corregidas.append((lat_corregida, lon_corregida))
    latitudes_decimal, longitudes_decimal = zip(*coordenadas_corregidas)
    df_copy["LATITUD_DECIMAL"], df_copy["LONGITUD_DECIMAL"] = latitudes_decimal, longitudes_decimal
    
    resultados = []
    for _, row in df_copy.iterrows():
        try:
            lat_raw, lon_raw = row["LATITUD_DECIMAL"], row["LONGITUD_DECIMAL"]
            if pd.isna(lat_raw) or pd.isna(lon_raw) or not (-90 <= lat_raw <= 90) or not (-180 <= lon_raw <= 180): continue
            distancia = geodesic((st.session_state.negocio_lat, st.session_state.negocio_lon), (lat_raw, lon_raw)).km
            if distancia < st.session_state.radio_km:
                tarifa_val = pd.to_numeric(row["TARIFA"], errors="coerce")
                if (st.session_state.presupuesto_min is not None and tarifa_val < st.session_state.presupuesto_min) or \
                   (st.session_state.presupuesto_max is not None and tarifa_val > st.session_state.presupuesto_max): continue
                
                maps_url = f"https://www.google.com/maps/place/{lat_raw},{lon_raw}"
                street_view_url = f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={lat_raw},{lon_raw}"

                resultados.append({
                    "CIUDAD": row.get("CIUDAD"), "CLAVE": row.get("CLAVE"), "DIRECCION": row.get("DIRECCION"),
                    "VISTA": row.get("VISTA"), "TIPO": row.get("TIPO"), "BASE": row.get("BASE"),
                    "ALTURA": row.get("ALTURA"), "AREA": row.get("AREA"), "LATITUD": lat_raw,
                    "LONGITUD": lon_raw, "DISTANCIA_KM": round(distancia, 2),
                    "TARIFA_PUBLICO": f"${tarifa_val:,.2f}", "IMPRESION": row.get("IMPRESION"),
                    "INSTALACION": row.get("INSTALACION"), "COSTO": row.get("IMPRESION+INSTALACION"),
                    "MAPS_": maps_url,
                    "STREET_VIEW": street_view_url,
                    "PROVEEDOR": row.get("PROVEEDOR"), "TELEFONO_PROVEEDOR": row.get("TELÉFONO PROVEEDOR"),
                })
        except Exception as e:
            st.warning(f"Error al procesar fila: {e}")

    if not resultados:
        st.warning("⚠️ No se encontraron espectaculares que cumplan con los criterios especificados.")
        st.session_state.df_filtrado = pd.DataFrame()
        st.session_state.busqueda_realizada = False
    else:
        st.session_state.df_filtrado = pd.DataFrame(resultados)
        st.success(f"✅ Búsqueda completada. Se encontraron **{len(resultados)}** resultados.")
        st.session_state.busqueda_realizada = True
        # Generar un nuevo folio solo cuando se realiza una nueva búsqueda exitosa
        st.session_state.folio_actual = generar_folio()

# 4. VISUALIZACIÓN Y DESCARGAS
if not st.session_state.df_filtrado.empty and st.session_state.busqueda_realizada:
    df_filtrado = st.session_state.df_filtrado
    st.write("---")
    st.header("🔍 **Resultados de la Búsqueda**")
    
    st.subheader("🗺️ Mapa de Espectaculares")
    mapa = folium.Map(location=[df_filtrado["LATITUD"].mean(), df_filtrado["LONGITUD"].mean()], zoom_start=13)
    folium.Marker(location=[st.session_state.negocio_lat, st.session_state.negocio_lon], popup=folium.Popup("<b>📍 Negocio</b>", max_width=300), icon=folium.Icon(color="red", icon="star")).add_to(mapa)
    folium.Circle(location=[st.session_state.negocio_lat, st.session_state.negocio_lon], radius=st.session_state.radio_km * 1000, color="blue", fill=True, fill_opacity=0.1).add_to(mapa)
    cluster = MarkerCluster().add_to(mapa)
    colores_marcadores = ["green", "orange", "purple", "cadetblue", "darkred", "darkgreen", "blue", "pink", "lightgreen", "black"]
    
    for i, r in df_filtrado.iterrows():
        try:
            color_marker = colores_marcadores[i % len(colores_marcadores)]
            popup_html = (f"<b>{r['CLAVE']}</b><b>{r['TIPO']}</b><br>Distancia: {r['DISTANCIA_KM']:.2f} km<br>Tarifa: {r['TARIFA_PUBLICO']}<br>"
                          f"<a href='{r['MAPS_']}' target='_blank'>📍 Google Maps</a><br>"
                          f"<a href='{r['STREET_VIEW']}' target='_blank'>🌐 Street View</a>")
            folium.Marker(
                location=[r["LATITUD"], r["LONGITUD"]],
                popup=folium.Popup(popup_html, max_width=300),
                icon=folium.Icon(color=color_marker, icon="info-sign"),
            ).add_to(cluster)
        except Exception as e:
            st.error(f"⚠️ No se pudo dibujar un marcador: {e}")
    st.components.v1.html(folium.Figure().add_child(mapa).render(), height=500)

    st.subheader("💾 Opciones de Descarga")
    col_dl1, col_dl2 = st.columns(2)
    
    # Mostrar el folio actual sin incrementarlo
    st.info(f"📋 **Folio de esta propuesta:** `{st.session_state.folio_actual}`")
    
    with col_dl1:
        csv_file = df_filtrado.to_csv(index=False).encode('utf-8')
        if st.download_button(
            label="⬇️ Descargar Resultados (CSV)", 
            data=csv_file, 
            file_name=f"{st.session_state.folio_actual}_resultados.csv", 
            mime="text/csv",
            key='download_csv'  # Clave única para este botón
        ):
            # Solo incrementar el contador después de la descarga exitosa
            incrementar_folio()
            st.success(f"✅ Descarga completada. Nuevo folio: `{generar_folio()}`")
    
    with col_dl2:
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "Lugares cercanos"
        ws.append(list(df_filtrado.columns))
        for r in df_filtrado.to_dict('records'):
            ws.append(list(r.values()))
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                if ws.cell(row=1, column=col).value == "TARIFA_PUBLICO":
                    ws.cell(row=row, column=col).number_format = '"$"#,##0.00'
                    break
        tabla = Table(displayName="TablaEspectaculares", ref=f"A1:{get_column_letter(ws.max_column)}{ws.max_row}")
        tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws.add_table(tabla)
        for cell in ws[1]:
            cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            cell.font = Font(bold=True)
        wb.save(output)
        if st.download_button(
            label="⬇️ Descargar Resultados (Excel)", 
            data=output.getvalue(), 
            file_name=f"{st.session_state.folio_actual}_resultados.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key='download_excel'  # Clave única para este botón
        ):
            # Solo incrementar el contador después de la descarga exitosa
            incrementar_folio()
            st.success(f"✅ Descarga completada. Nuevo folio: `{generar_folio()}`")
    
    # 5. GENERAR PRESENTACIÓN
    st.write("---")
    st.subheader("🎓 Generar Presentación")
    st.write("Selecciona los espectaculares que deseas incluir en una presentación PowerPoint.")
    opciones_presentacion = [f"{i+1}. {r['CLAVE']} - {r['TARIFA_PUBLICO']}" for i, r in df_filtrado.iterrows()]
    seleccionados_presentacion = st.multiselect("Elige los espectaculares de la lista:", opciones_presentacion, placeholder="Selecciona 1 o más...")
    
    if st.button("Crear Presentación", key='crear_presentacion'):
        if not seleccionados_presentacion:
            st.warning("⚠️ Por favor, selecciona al menos un espectacular para crear la presentación.")
        else:
            indices_seleccionados = [int(op.split(".")[0]) - 1 for op in seleccionados_presentacion]
            seleccionados_df = df_filtrado.iloc[indices_seleccionados]
            try:
                plantilla_pptx = "plantilla2.pptx"
                prs = Presentation(plantilla_pptx)
                if len(prs.slides) < 2:
                    st.error("❌ La plantilla debe tener al menos 2 diapositivas: la de título (0) y la de contenido (1).")
                else:
                    slide_base_index = 1
                    slide_base = prs.slides[slide_base_index]
                    for _, fila in seleccionados_df.iterrows():
                        nueva_slide = duplicar_slide(prs, slide_base)
                        reemplazar_texto_slide(nueva_slide, fila.to_dict(), seleccionados_df)
                    if not seleccionados_df.empty:
                        primera_fila = seleccionados_df.iloc[0]
                        reemplazar_texto_slide(prs.slides[0], primera_fila.to_dict(), seleccionados_df)
                    pptx_output = io.BytesIO()
                    prs.save(pptx_output)
                    
                    st.success(f"✅ ¡Presentación creada con éxito! - Folio: `{st.session_state.folio_actual}`")
                    
                    if st.download_button(
                        label="⬇️ Descargar Presentación (PPTX)", 
                        data=pptx_output.getvalue(), 
                        file_name=f"{st.session_state.folio_actual}.pptx", 
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key='download_pptx'  # Clave única para este botón
                    ):
                        # Solo incrementar el contador después de la descarga exitosa
                        incrementar_folio()
                        st.success(f"✅ Descarga completada. Nuevo folio: `{generar_folio()}`")
            except FileNotFoundError:
                st.error("❌ **Error:** No se encontró el archivo de plantilla `plantilla2.pptx`. Asegúrate de que está en la misma carpeta que tu `app.py`.")
            except Exception as e:
                st.error(f"❌ **Error al crear la presentación:** {e}")


