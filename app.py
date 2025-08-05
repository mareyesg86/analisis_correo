import streamlit as st
import pandas as pd
from collections import Counter
from textblob import TextBlob
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
import plotly.express as px
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from email.utils import parseaddr

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

archivo = st.sidebar.file_uploader("📤 Sube tu archivo .CSV exportado desde Outlook", type="csv")

if archivo:
    df = pd.read_csv(archivo, encoding='latin1')
    df.columns = df.columns.str.strip()
    df['De'] = df['De'].str.strip()
    df['Nombre'] = df['De'].apply(lambda x: parseaddr(x)[0].strip() or parseaddr(x)[1].split('@')[0])
    df['Email'] = df['De'].apply(lambda x: parseaddr(x)[1].lower())
    df['Asunto'] = df['Asunto'].astype(str).str.strip().str.lower()
    df['Cuerpo'] = df['Cuerpo'].astype(str).str.strip().str.lower()
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce', dayfirst=False)
    df = df.dropna(subset=['Fecha'])

    palabras_clave = ['urgente', 'favor', 'reunión', 'entrega', 'plazo', 'respuesta', 'pendiente', 'informe', 'revisar']

    def detectar_keywords(texto):
        return list(set([p for p in palabras_clave if p in texto]))

    # Sentimiento TextBlob
    def sentimiento_textblob(texto):
        blob = TextBlob(texto)
        p = blob.sentiment.polarity
        if p > 0.2:
            return 'Positivo'
        elif p < -0.2:
            return 'Negativo'
        else:
            return 'Neutro'

    # Sentimiento VADER
    analyzer = SentimentIntensityAnalyzer()
    def sentimiento_vader(texto):
        score = analyzer.polarity_scores(texto)
        if score['compound'] > 0.2:
            return 'Positivo'
        elif score['compound'] < -0.2:
            return 'Negativo'
        else:
            return 'Neutro'

    def clasificacion_accion(texto):
        texto = texto.lower()
        if any(p in texto for p in ["urgente", "favor", "responder", "pendiente"]):
            return "🔴 Leer y responder"
        elif any(p in texto for p in ["delegar", "avisar a", "encargarse"]):
            return "🟡 Delegar"
        elif any(p in texto for p in ["solo informar", "fyi", "adjunto informe", "revisión"]):
            return "🟢 Puede esperar"
        elif any(p in texto for p in ["entregar informe", "realizar", "subir a plataforma"]):
            return "🔵 Tarea propia"
        else:
            return "🟢 Puede esperar"

    resumen = []
    for (nombre, email), grupo in df.groupby(['Nombre', 'Email']):
        total = len(grupo)
        palabras = []
        sentimientos_tb = []
        sentimientos_vd = []
        acciones = []

        for _, fila in grupo.iterrows():
            texto = f"{fila['Asunto']} {fila['Cuerpo']}"
            palabras += detectar_keywords(texto)
            sentimientos_tb.append(sentimiento_textblob(texto))
            sentimientos_vd.append(sentimiento_vader(texto))
            acciones.append(clasificacion_accion(texto))

        palabras_frecuentes = Counter(palabras)
        sentimiento_tb_dom = Counter(sentimientos_tb).most_common(1)[0][0]
        sentimiento_vd_dom = Counter(sentimientos_vd).most_common(1)[0][0]
        accion_predominante = Counter(acciones).most_common(1)[0][0]
        prioritario = 'Sí' if 'urgente' in palabras_frecuentes or 'entrega' in palabras_frecuentes or sentimiento_tb_dom == 'Negativo' else 'No'

        resumen.append({
            'Nombre Remitente': nombre,
            'Correo': email,
            'Correos Recibidos': total,
            'Palabras Clave': ', '.join(palabras_frecuentes.keys()) if palabras_frecuentes else 'Ninguna',
            'Sentimiento TextBlob': sentimiento_tb_dom,
            'Sentimiento VADER': sentimiento_vd_dom,
            '¿Priorizar Respuesta?': prioritario,
            'Clasificación de Acción': accion_predominante
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
        por_remitente = df['Nombre'].value_counts().nlargest(10).reset_index()
        por_remitente.columns = ['Remitente', 'Cantidad']
        fig2 = px.bar(por_remitente, x='Remitente', y='Cantidad', title='Top 10 Remitentes')
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### ☁️ Nube de palabras (corpus completo)")
    texto_completo = " ".join(df['Cuerpo'].astype(str))
    if texto_completo.strip():
        wordcloud = WordCloud(width=800, height=400, background_color="white", colormap='viridis').generate(texto_completo)
        fig_wc, ax_wc = plt.subplots(figsize=(10, 5))
        ax_wc.imshow(wordcloud, interpolation="bilinear")
        ax_wc.axis("off")
        st.pyplot(fig_wc)
    else:
        st.info("No hay suficiente texto para generar la nube de palabras.")

    st.markdown("### 🧠 Análisis por Remitente")

    filtro_prioridad = st.selectbox("Filtrar por prioridad:", ["Todos", "Solo prioritarios"])
    buscar_remitente = st.text_input("Buscar remitente contiene:", "").lower()

    df_filtrado = resumen_df.copy()
    if filtro_prioridad == "Solo prioritarios":
        df_filtrado = df_filtrado[df_filtrado['¿Priorizar Respuesta?'] == 'Sí']
    if buscar_remitente:
        df_filtrado = df_filtrado[df_filtrado['Nombre Remitente'].str.lower().str.contains(buscar_remitente)]

    st.dataframe(df_filtrado, use_container_width=True)

    csv_export = df_filtrado.to_csv(index=False).encode('utf-8-sig')
    st.download_button("📁 Descargar resumen filtrado como CSV", csv_export, "resumen_correos.csv", "text/csv")

else:
    st.warning("📂 Sube un archivo .csv exportado desde Outlook para comenzar.")
