import streamlit as st
import re
import io
from docx import Document

# Importaciones de la librería (compatibles con v1.x)
from youtube_transcript_api import YouTubeTranscriptApi
try:
    from youtube_transcript_api import (
        TranscriptsDisabled,
        NoTranscriptFound,
        VideoUnavailable,
        CouldNotRetrieveTranscript,
    )
except ImportError:
    # Fallback para versiones que aún exponen los errores en _errors
    from youtube_transcript_api._errors import (
        TranscriptsDisabled,
        NoTranscriptFound,
        VideoUnavailable,
        CouldNotRetrieveTranscript,
    )

# ---------------------------------------------------------------
# Configuración de la página
# ---------------------------------------------------------------
st.set_page_config(
    page_title="YouTube a Word",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Transcripción de YouTube a Word")
st.caption(
    "Pega la URL de un video de YouTube para descargar su transcripción en formato .docx"
)

# ---------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------
def extraer_id_video(url: str):
    """Extrae el ID de 11 caracteres de una URL de YouTube."""
    patrones = [
        r'(?:youtube\.com\/watch\?v=)([a-zA-Z0-9_-]{11})',
        r'(?:youtu\.be\/)([a-zA-Z0-9_-]{11})',
        r'(?:youtube\.com\/embed\/)([a-zA-Z0-9_-]{11})',
        r'(?:youtube\.com\/shorts\/)([a-zA-Z0-9_-]{11})',
    ]
    for patron in patrones:
        match = re.search(patron, url)
        if match:
            return match.group(1)
    return None


def generar_docx(texto: str, url_video: str) -> bytes:
    """Crea un documento de Word en memoria y devuelve sus bytes."""
    doc = Document()
    doc.add_heading("Transcripción del Video de YouTube", 0)
    doc.add_paragraph(f"URL original: {url_video}")
    doc.add_paragraph("")  # Línea en blanco

    # Dividir en párrafos para mejor legibilidad
    parrafos = re.split(r'(?<=[.!?])\s+', texto)
    for parrafo in parrafos:
        if parrafo.strip():
            doc.add_paragraph(parrafo.strip())

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def obtener_transcripcion(video_id: str):
    """Obtiene la transcripción usando la API v1.x de youtube-transcript-api."""
    try:
        # 1) Instanciar la API (nueva forma desde v1.0)
        ytt_api = YouTubeTranscriptApi()

        # 2) Listar transcripciones disponibles
        lista = ytt_api.list(video_id)

        # 3) Estrategia de selección: español > inglés > cualquiera
        transcript = None
        for idioma in ("es", "en"):
            try:
                try:
                    transcript = lista.find_manually_created_transcript([idioma])
                except NoTranscriptFound:
                    transcript = lista.find_generated_transcript([idioma])
                break  # encontrada, salir del bucle
            except NoTranscriptFound:
                continue

        if transcript is None:
            transcript = list(lista)[0]

        # 4) Obtener los fragmentos
        datos = transcript.fetch()

        # 5) En v1.x los snippets son objetos con atributo .text
        texto = " ".join(snippet.text for snippet in datos)
        return texto

    except TranscriptsDisabled:
        st.error("❌ El video tiene los subtítulos desactivados o no dispone de ellos.")
    except NoTranscriptFound:
        st.error("❌ No se encontró ninguna transcripción para este video.")
    except VideoUnavailable:
        st.error("❌ El video no está disponible (privado o eliminado).")
    except CouldNotRetrieveTranscript as e:
        st.error(f"❌ No se pudo recuperar la transcripción. Detalle: {e}")
    except Exception as e:
        st.error(f"❌ Ocurrió un error inesperado: {e}")

    return None


# ---------------------------------------------------------------
# Interfaz de usuario
# ---------------------------------------------------------------
with st.form("url_form"):
    url_input = st.text_input(
        "🔗 URL del video de YouTube",
        placeholder="https://www.youtube.com/watch?v=...",
        help="También funciona con enlaces cortos (youtu.be) o Shorts.",
    )
    submitted = st.form_submit_button("Obtener transcripción")

if submitted:
    if not url_input.strip():
        st.warning("⚠️ Por favor, ingresa una URL.")
    else:
        video_id = extraer_id_video(url_input)
        if not video_id:
            st.warning(
                "⚠️ No se pudo identificar un ID de video válido en la URL. Verifícala."
            )
        else:
            with st.spinner("Buscando y procesando la transcripción..."):
                texto = obtener_transcripcion(video_id)

            if texto:
                st.success("✅ ¡Transcripción obtenida con éxito!")

                docx_bytes = generar_docx(texto, url_input)

                st.download_button(
                    label="📥 Descargar archivo Word (.docx)",
                    data=docx_bytes,
                    file_name=f"transcripcion_{video_id}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument"
                        ".wordprocessingml.document"
                    ),
                    type="primary",
                )

                with st.expander("👀 Vista previa del texto"):
                    preview = texto[:1000] + "..." if len(texto) > 1000 else texto
                    st.write(preview)
