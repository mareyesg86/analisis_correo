import streamlit as st
import pandas as pd
from collections import Counter
from textblob import TextBlob
import plotly.express as px
from datetime import datetime
import google.generativeai as genai
import os

st.set_page_config(page_title="Dashboard Correos Outlook", layout="wide")

st.markdown("""
<style>
.big-title {
    font-size: 2.5em;
    font-weight: 700;
    color: #2c3e50;
}
.metric-label {
    font-size: 0.9em;
    color: #888;
}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='big-title'>📬 Dashboard de Correos Outlook</div>", unsafe_allow_html=True)

# Configurar clave API de Gemini (desde Secrets o input seguro)
gemini_api_key = st.secrets["GEMINI_API_KEY"] if "GEMINI_API_KEY" in st.secrets else st.text_input("🔑 API Key de Gemini", type="password")

if gemini_api_key:
    genai.configure(api_key=gemini_api_key)
    model = genai.GenerativeModel("models/gemini-1.5-flash-latest")  # usa ruta oficial

archivo = st.sidebar.file_uploader("📤 Sube tu archivo .CSV exportado desde Outlook", type="csv")

if archivo:
    df = pd.read_csv(archivo, encoding='latin1')
    df.columns = df.columns.str.strip()
    df['De'] = df['De'].str.strip().str.lower()
    df['Asunto'] = df['Asunto'].str.strip().str.lower()
    df['Cuerpo'] = df['Cuerpo'].str.strip().str.lower()
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce', dayfirst=False)
    df = df.dropna(subset=['Fecha'])

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

    def generar_resumen(texto):
        if not gemini_api_key:
            return "Gemini API no configurada."
        prompt = f"Resume brevemente el siguiente correo en 1 frase clara y directa:\n\n{texto}"
        try:
            response = model.generate_content(prompt)
            return response.text.strip()
        except Exception as e:
            return f"Error: {e}"

    resumen = []
    for remitente, grupo in df.groupby('De'):
        total = len(grupo)
        palabras = []
        sentimientos = []
        resumen_textos = []

        for _, fila in grupo.iterrows():
            texto = f"{fila['Asunto']} {fila['Cuerpo']}"
            palabras += detectar_keywords(texto)
            sentimientos.append(sentimiento(texto))
            if gemini_api_key:
                resumen_textos.append(generar_resumen(texto))

        palabras_frecuentes = Counter(palabras)
        sentimiento_dominante = Counter(sentimientos).most_common(1)[0][0]
        prioritario = 'Sí' if 'urgente' in palabras_frecuentes or 'entrega' in palabras_frecuentes or sentimiento_dominante == 'Negativo' else 'No'

        resumen.append({
            'Remitente': remitente,
            'Correos Recibidos': total,
            'Palabras Clave': ', '.join(palabras_frecuentes.keys()) if palabras_frecuentes else 'Ninguna',
            'Sentimiento Dominante': sentimiento_dominante,
            '¿Priorizar Respuesta?': prioritario,
            'Resumen IA': resumen_textos[0] if resumen_textos else ''
        })

    resumen_df = pd.DataFrame(resumen).sort_values(by='Correos Recibidos', ascending=False)

    st.markdown("### 📈 Métricas Generales")
    col1, col2, col3 = st.columns(3)
    col1.metric("Correos Totales", len(df))
    col2.metric("Remitentes Únicos", resumen_df.shape[0])
    col3.metric("Correos Prioritarios", resumen_df[resumen_df['¿Priorizar Respuesta?'] == 'Sí'].shape[0])

    st.markdown("### 📊 Visualizaciones")
    col1, col2 = st.columns(2)

    with col1:
        por_dia = df.groupby(df['Fecha'].dt.date).size().reset_index(name='Cantidad')
        fig1 = px.bar(por_dia, x='Fecha', y='Cantidad', title='Correos por Día', labels={'Cantidad': 'N° de Correos'})
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        por_remitente = df['De'].value_counts().nlargest(10).reset_index()
        por_remitente.columns = ['Remitente', 'Cantidad']
        fig2 = px.bar(por_remitente, x='Remitente', y='Cantidad', title='Top 10 Remitentes')
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### 🧠 Análisis por Remitente")

    filtro_prioridad = st.selectbox("Filtrar por prioridad:", ["Todos", "Solo prioritarios"])
    buscar_remitente = st.text_input("Buscar remitente contiene:", "").lower()

    df_filtrado = resumen_df.copy()
    if filtro_prioridad == "Solo prioritarios":
        df_filtrado = df_filtrado[df_filtrado['¿Priorizar Respuesta?'] == 'Sí']
    if buscar_remitente:
        df_filtrado = df_filtrado[df_filtrado['Remitente'].str.contains(buscar_remitente)]

    st.dataframe(df_filtrado, use_container_width=True)

    csv_export = df_filtrado.to_csv(index=False).encode('utf-8-sig')
    st.download_button("📁 Descargar resumen filtrado como CSV", csv_export, "resumen_correos.csv", "text/csv")

else:
    st.warning("📂 Sube un archivo .csv exportado desde Outlook para comenzar.")
