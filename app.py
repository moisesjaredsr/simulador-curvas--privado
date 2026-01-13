import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import xlsxwriter

# --- OCULTAR MENÚS Y MARCAS DE AGUA (SEGURIDAD VISUAL) ---
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)


# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Simulador Solar Web", layout="wide", page_icon="☀️")

st.title("☀️ Simulador de Análisis de Celdas Solares")
st.markdown("---")

# --- BARRA LATERAL (CONTROLES) ---
with st.sidebar:
    st.header("1. Parámetros Físicos")
    area = st.number_input("Área (cm²)", value=0.121, step=0.001, format="%.3f")
    potencia = st.number_input("Potencia (W/cm²)", value=0.1, step=0.01)

    st.header("2. Límites de Gráfica")
    col1, col2 = st.columns(2)
    xmin = col1.number_input("X Min", value=0.0)
    xmax = col2.number_input("X Max", value=0.8)
    ymin = col1.number_input("Y Min", value=-5.0)
    ymax = col2.number_input("Y Max", value=20.0)

# --- ZONA DE CARGA DE ARCHIVOS ---
st.subheader("Cargar Archivos (.txt)")
uploaded_files = st.file_uploader("Arrastra tus archivos aquí", type=["txt"], accept_multiple_files=True)

if uploaded_files:
    # Listas para guardar resultados
    resultados_lista = []
    datos_para_excel = {}
    
    # Crear la figura interactiva
    fig = go.Figure()

    # --- EL CORAZÓN (TUS CÁLCULOS) ---
    for uploaded_file in uploaded_files:
        try:
            # Leer el archivo desde la memoria (Streamlit)
            df = pd.read_csv(uploaded_file, sep=r'\s+', skiprows=1, header=None, engine='python')
            df = df.dropna()
            
            # Limpieza de datos (Coma por punto)
            df[0] = df[0].astype(str).str.replace(',', '.').astype(float)
            df[1] = df[1].astype(str).str.replace(',', '.').astype(float)

            # --- TUS MATEMÁTICAS EXACTAS ---
            VF = 1 * df[1].values
            IM_raw = -1 * df[0].values / area
            IM_mA = IM_raw * 1000

            # Cálculos de eficiencia
            P = IM_raw * VF
            P_max = np.max(P)
            Eta = (P_max / potencia) * 100

            # Jsc (Corriente en V=0)
            idx_jsc = (np.abs(VF - 0)).argmin()
            val_Jsc = IM_mA[idx_jsc]

            # Voc (Voltaje en I=0)
            idx_voc = (np.abs(IM_mA - 0)).argmin()
            val_Voc = VF[idx_voc]

            # Fill Factor
            Jsc_amp = val_Jsc / 1000.0
            if Jsc_amp * val_Voc != 0:
                val_FF = 100 * (P_max / (Jsc_amp * val_Voc))
            else:
                val_FF = 0

            nombre_archivo = uploaded_file.name

            # Guardar resultados
            resultados_lista.append({
                "Archivo": nombre_archivo,
                "Jsc (mA)": round(val_Jsc, 4),
                "Voc (V)": round(val_Voc, 4),
                "FF (%)": round(val_FF, 2),
                "Eta (%)": round(Eta, 2)
            })

            # Guardar datos crudos para Excel
            datos_para_excel[nombre_archivo] = {
                'Voltaje': VF,
                'Corriente': IM_mA
            }

            # Añadir curva a la gráfica interactiva
            fig.add_trace(go.Scatter(x=VF, y=IM_mA, mode='lines', name=nombre_archivo))

        except Exception as e:
            st.error(f"Error procesando {uploaded_file.name}: {e}")

    # --- MOSTRAR RESULTADOS ---
    
    # 1. Tabla de Datos
    st.subheader("📊 Tabla de Resultados")
    df_resultados = pd.DataFrame(resultados_lista)
    st.dataframe(df_resultados, use_container_width=True)

    # 2. Gráfica Interactiva
    st.subheader("📈 Curva IV")
    fig.update_layout(
        xaxis_title="Voltaje (V)",
        yaxis_title="Densidad de Corriente (mA/cm²)",
        xaxis_range=[xmin, xmax],
        yaxis_range=[ymin, ymax],
        template="plotly_white",
        height=600
    )
    # Líneas de ejes cero
    fig.add_hline(y=0, line_width=1, line_color="black")
    fig.add_vline(x=0, line_width=1, line_color="black")
    
    st.plotly_chart(fig, use_container_width=True)

    # --- EXPORTAR A EXCEL ---
    st.subheader("💾 Exportar Reporte")
    
    # Generar Excel en memoria
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Hoja 1: Resultados
        df_resultados.to_excel(writer, sheet_name='Resultados', index=False)
        worksheet = writer.sheets['Resultados']
        worksheet.set_column('A:A', 25)
        
        # Hoja 2: Datos Crudos
        df_crudos = pd.DataFrame()
        for nombre, datos in datos_para_excel.items():
            df_crudos[f"V_{nombre}"] = pd.Series(datos['Voltaje'])
            df_crudos[f"J_{nombre}"] = pd.Series(datos['Corriente'])
        
        df_crudos.to_excel(writer, sheet_name='Datos_Curvas', index=False)
        
    
    # Botón de Descarga
    st.download_button(
        label="Descargar Excel con Resultados",
        data=buffer.getvalue(),
        file_name="Reporte_Solar.xlsx",
        mime="application/vnd.ms-excel"
    )

else:
    st.info("👆 Sube tus archivos .txt usando el panel de arriba para comenzar.")
