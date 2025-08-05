import streamlit as st
import pandas as pd
from collections import Counter
from textblob import TextBlob
import plotly.express as px
from datetime import datetime
import io

st.set_page_config(page_title="Análisis de Correos Outlook", layout="wide")
st.title("📧 Análisis de Correos Exportados desde Outlook")

# --- Subida de archivo CSV ---
archivo = st.file_uploader("Sube el archivo .CSV exportado desde Outlook (Fecha, De, Asunto, Cuerpo)", type="csv")

if archivo:
    df = pd.read_csv(archivo, encoding='latin1')
    df.columns = df.columns.str.strip()
    df['De'] = df['De'].str.strip().str.lower()
    df['Asunto'] = df['Asunto'].str.strip().str.lower()
    df['Cuerpo'] = df['Cuerpo'].str.strip().str.lower()

    # Convertir fechas
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce', dayfirst=False)
    df = df.dropna(subset=['Fecha'])

    st.subheader("Vista previa del archivo cargado")
    st.dataframe(df.head(), use_container_width=True)

    # --- Análisis ---
    palabras_clave = ['urgente', 'favor', 'reunión', 'entrega', 'plazo', 'respuesta', 'pendiente', 'informe', 'revisar']

    def detectar_keywords(texto):
        return list(set([p for p in palabras_clave if p in texto]))

    def sentimiento(texto):
        blob = TextBlob(texto)
        p = blob.sentiment.polarity
        if p > 0.2:
            return 'Positivo'
        elif p < -0.2:
            return 'Negativo'
        else:
            return 'Neutro'

    resumen = []
    for remitente, grupo in df.groupby('De'):
        total = len(grupo)
        palabras = []
        sentimientos = []

        for _, fila in grupo.iterrows():
            texto = f"{fila['Asunto']} {fila['Cuerpo']}"
            palabras += detectar_keywords(texto)
            sentimientos.append(sentimiento(texto))

        palabras_frecuentes = Counter(palabras)
        sentimiento_dominante = Counter(sentimientos).most_common(1)[0][0]
        prioritario = 'Sí' if 'urgente' in palabras_frecuentes or 'entrega' in palabras_frecuentes or sentimiento_dominante == 'Negativo' else 'No'

        resumen.append({
            'Remitente': remitente,
            'Correos Recibidos': total,
            'Palabras Clave': ', '.join(palabras_frecuentes.keys()) if palabras_frecuentes else 'Ninguna',
            'Sentimiento Dominante': sentimiento_dominante,
            '¿Priorizar Respuesta?': prioritario
        })

    resumen_df = pd.DataFrame(resumen).sort_values(by='Correos Recibidos', ascending=False)

    st.subheader("🔢 Análisis Agrupado por Remitente")
    st.dataframe(resumen_df, use_container_width=True)

    # Filtros
    st.sidebar.title("🔍 Filtros")
    filtro_prioridad = st.sidebar.selectbox("Mostrar solo...", ["Todos", "Solo prioritarios"])
    buscar_remitente = st.sidebar.text_input("Filtrar remitente contiene...").lower()

    df_filtrado = resumen_df.copy()
    if filtro_prioridad == "Solo prioritarios":
        df_filtrado = df_filtrado[df_filtrado['¿Priorizar Respuesta?'] == 'Sí']
    if buscar_remitente:
        df_filtrado = df_filtrado[df_filtrado['Remitente'].str.contains(buscar_remitente)]

    st.dataframe(df_filtrado, use_container_width=True)

    # --- Gráficos ---
    st.subheader("📊 Visualización de Actividad")
    col1, col2 = st.columns(2)

    with col1:
        por_dia = df.groupby(df['Fecha'].dt.date).size().reset_index(name='Cantidad')
        fig1 = px.bar(por_dia, x='Fecha', y='Cantidad', title='Correos recibidos por día')
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        por_remitente = df['De'].value_counts().nlargest(10).reset_index()
        por_remitente.columns = ['Remitente', 'Cantidad']
        fig2 = px.bar(por_remitente, x='Remitente', y='Cantidad', title='Top 10 remitentes')
        st.plotly_chart(fig2, use_container_width=True)

    # --- Exportar CSV ---
    csv_export = resumen_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button("📁 Descargar resumen como CSV", csv_export, "resumen_correos.csv", "text/csv")

else:
    st.info("Por favor, sube un archivo CSV para comenzar.")
