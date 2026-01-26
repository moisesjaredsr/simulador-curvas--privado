import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import xlsxwriter

# --- ESTILOS VISUALES ---
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Simulador Solar Web", layout="wide", page_icon="☀️")
st.title("☀️ Simulador de Análisis de Celdas (Final)")
st.markdown("---")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Configuración de Medición")
    
    # Checkbox principal
    is_dark = st.checkbox("¿Es medición en OSCURIDAD?", value=True)
    
    # Umbral para oscuridad
    turn_on_threshold = 0.0
    if is_dark:
        st.info("Define 'comienza a aumentar':")
        turn_on_threshold = st.number_input(
            "Corriente de corte (mA/cm²)", 
            value=0.1, 
            step=0.01, 
            format="%.3f",
            help="El voltaje se calculará exactamente cuando la corriente cruce este valor."
        )

    st.header("2. Parámetros Físicos")
    area = st.number_input("Área (cm²)", value=0.121, step=0.001, format="%.3f")
    potencia = st.number_input("Potencia (W/cm²)", value=0.1, step=0.01)

    st.header("3. Límites de Gráfica (Visualización Web)")
    col1, col2 = st.columns(2)
    xmin = col1.number_input("X Min", value=-0.1) 
    xmax = col2.number_input("X Max", value=1.5)
    ymin = col1.number_input("Y Min", value=-1.0)
    ymax = col2.number_input("Y Max", value=20.0)

# --- FUNCIÓN DE INTERPOLACIÓN ---
def calcular_interseccion(x, y, target_y=0):
    """Calcula x cuando y = target_y usando interpolación lineal."""
    x = np.array(x)
    y = np.array(y)
    
    # Ordenar por X
    idx_sort = np.argsort(x)
    x = x[idx_sort]
    y = y[idx_sort]

    y_diff = y - target_y
    sign_changes = np.where(np.diff(np.signbit(y_diff)))[0]
    
    if len(sign_changes) > 0:
        idx = sign_changes[0] 
        x1, x2 = x[idx], x[idx+1]
        y1, y2 = y[idx], y[idx+1]
        
        if y1 == y2: return x1
        return x1 + (target_y - y1) * (x2 - x1) / (y2 - y1)
    else:
        return x[np.abs(y_diff).argmin()]

# --- PROCESAMIENTO ---
st.subheader("Cargar Archivos (.txt)")
uploaded_files = st.file_uploader("Arrastra tus archivos aquí", type=["txt"], accept_multiple_files=True)

if uploaded_files:
    resultados_lista = []
    datos_para_excel = {}
    fig = go.Figure()

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_csv(uploaded_file, sep=r'\s+', skiprows=1, header=None, engine='python')
            df = df.dropna()
            
            # Limpieza
            df[0] = df[0].astype(str).str.replace(',', '.').astype(float)
            df[1] = df[1].astype(str).str.replace(',', '.').astype(float)

            # Variables
            VF = df[1].values
            
            # Ajuste de Polaridad
            factor_polaridad = 1.0 if is_dark else -1.0
            IM_raw = factor_polaridad * df[0].values / area
            IM_mA = IM_raw * 1000

            # Ordenar (Vital para interpolación y Excel)
            sort_idx = np.argsort(VF)
            VF = VF[sort_idx]
            IM_mA = IM_mA[sort_idx]
            IM_raw = IM_raw[sort_idx]

            # --- CÁLCULOS SEGÚN MODO ---
            if is_dark:
                # MODO OSCURIDAD
                val_Voc = calcular_interseccion(x=VF, y=IM_mA, target_y=turn_on_threshold)
                
                val_Jsc_str = "N/A"
                val_FF_str = "N/A"
                val_Eta_str = "N/A"
                label_voc = f"V_turn-on (@{turn_on_threshold}mA)" 
                
            else:
                # MODO LUZ
                val_Voc = calcular_interseccion(x=VF, y=IM_mA, target_y=0)
                val_Jsc = calcular_interseccion(x=IM_mA, y=VF, target_y=0)
                
                P = IM_raw * VF
                P_max = np.max(P)
                Eta = (P_max / potencia) * 100
                
                Jsc_amp = val_Jsc / 1000.0
                if Jsc_amp * val_Voc != 0:
                    val_FF = 100 * (P_max / (Jsc_amp * val_Voc))
                else:
                    val_FF = 0
                    
                val_Jsc_str = round(val_Jsc, 4)
                val_FF_str = round(val_FF, 2)
                val_Eta_str = round(Eta, 2)
                label_voc = "Voc (V)"

            # Guardar en lista
            resultados_lista.append({
                "Archivo": uploaded_file.name,
                label_voc: round(val_Voc, 4),
                "Jsc (mA)": val_Jsc_str,
                "FF (%)": val_FF_str,
                "Eta (%)": val_Eta_str
            })

            # Guardar datos para excel
            datos_para_excel[uploaded_file.name] = {'Voltaje': VF, 'Corriente': IM_mA}

            # Gráfica Web
            fig.add_trace(go.Scatter(x=VF, y=IM_mA, mode='lines', name=uploaded_file.name))

        except Exception as e:
            st.error(f"Error en {uploaded_file.name}: {e}")

    # --- MOSTRAR RESULTADOS EN PANTALLA ---
    st.subheader("📊 Tabla de Resultados")
    df_res = pd.DataFrame(resultados_lista)
    st.dataframe(df_res, use_container_width=True)

    st.subheader("📈 Curva I-V")
    y_label_web = "Corriente de Diodo (mA/cm²)" if is_dark else "Fotocorriente (mA/cm²)"
    
    fig.update_layout(
        xaxis_title="Voltaje (V)",
        yaxis_title=y_label_web,
        xaxis_range=[xmin, xmax],
        yaxis_range=[ymin, ymax],
        template="plotly_white",
        height=600
    )
    
    if is_dark:
        fig.add_hline(y=turn_on_threshold, line_dash="dot", line_color="red", annotation_text="Umbral")
    else:
        fig.add_hline(y=0, line_color="black", line_width=1)
    
    fig.add_vline(x=0, line_color="black", line_width=1)
    st.plotly_chart(fig, use_container_width=True)

    # --- EXPORTAR A EXCEL (CON GRÁFICA EDITABLE) ---
    st.subheader("💾 Exportar Excel con Gráfica")
    
    buffer = io.BytesIO()
    
    # Preparamos el DataFrame de datos crudos antes de abrir el writer
    df_raw = pd.DataFrame()
    for k, v in datos_para_excel.items():
        df_raw[f"V_{k}"] = pd.Series(v['Voltaje'])
        df_raw[f"I_{k}"] = pd.Series(v['Corriente'])
        
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # 1. Escribir Hoja de Resultados
        df_res.to_excel(writer, sheet_name='Resultados', index=False)
        worksheet_res = writer.sheets['Resultados']
        worksheet_res.set_column('A:E', 15)
        
        # 2. Escribir Hoja de Datos Crudos
        df_raw.to_excel(writer, sheet_name='Datos_Crudos', index=False)
        
        # 3. CREACIÓN DE LA GRÁFICA NATIVA EN EXCEL
        workbook = writer.book
        
        # Crear objeto gráfico (Scatter con líneas suaves)
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        
        # Definir rango de datos. 
        # La hoja 'Datos_Crudos' tiene columnas alternas: V (par), I (impar)
        # 0=A, 1=B, 2=C, 3=D...
        num_filas = len(df_raw)
        
        # Iterar sobre cada archivo cargado para agregar su serie a la gráfica
        for i, nombre_archivo in enumerate(datos_para_excel.keys()):
            col_v = i * 2       # Columnas 0, 2, 4...
            col_i = i * 2 + 1   # Columnas 1, 3, 5...
            
            chart.add_series({
                'name':       nombre_archivo,
                'categories': ['Datos_Crudos', 1, col_v, num_filas, col_v], # Eje X: Voltaje
                'values':     ['Datos_Crudos', 1, col_i, num_filas, col_i], # Eje Y: Corriente
                'line':       {'width': 1.5},
            })
            
        # Configurar títulos y ejes del gráfico Excel
        chart.set_title ({'name': 'Curvas I-V'})
        chart.set_x_axis({'name': 'Voltaje (V)', 'major_gridlines': {'visible': True}})
        chart.set_y_axis({'name': y_label_web,   'major_gridlines': {'visible': True}})
        
        # Insertar la gráfica en la hoja 'Resultados' (a la derecha de la tabla)
        worksheet_res.insert_chart('G2', chart, {'x_scale': 2, 'y_scale': 2}) 
        # x_scale y y_scale hacen la gráfica más grande

    st.download_button(
        label="Descargar Reporte Completo (Resultados + Gráfica)",
        data=buffer.getvalue(),
        file_name="Reporte_Solar_Completo.xlsx",
        mime="application/vnd.ms-excel"
    )

else:
    st.info("Carga tus archivos .txt arriba.")
