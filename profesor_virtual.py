import streamlit as st
from openai import OpenAI
from docx import Document
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet


# --------------------------------
# CONFIGURACIÓN
# --------------------------------

st.set_page_config(
    page_title="Profesor Virtual",
    page_icon="🧬"
)

try:
    client = OpenAI(
        api_key=st.secrets["OPENAI_API_KEY"]
    )
except:
    client = OpenAI()

# ID del Vector Store
VECTOR_STORE_ID = "vs_6a9ef1bd8a1c8191a51660152db783ed"


# --------------------------------
# FUNCIÓN PARA CREAR EL WORD
# --------------------------------

def crear_word(texto, tipo_tarea, idioma):
    documento = Document()

    # Título
    documento.add_heading("Profesor Virtual", level=1)

    # Información del documento
    documento.add_paragraph(
        "Biología y Geología · 1.º Bachillerato"
    )
    documento.add_paragraph(
        f"Tipo de tarea: {tipo_tarea}"
    )
    documento.add_paragraph(
        f"Idioma: {idioma}"
    )

    documento.add_paragraph("")

    # Contenido
    documento.add_heading("Contenido", level=2)

    for linea in texto.split("\n"):
        if linea.strip():
            documento.add_paragraph(linea)

    # Crear el archivo Word en memoria
    archivo = BytesIO()
    documento.save(archivo)
    archivo.seek(0)

    return archivo


# --------------------------------
# INTERFAZ
# --------------------------------

st.title("🧬 Profesor Virtual")
st.subheader("Biología y Geología · 1.º Bachillerato")

st.write(
    "Haz una pregunta o solicita una actividad "
    "sobre los contenidos de la asignatura."
)

idioma = st.selectbox(
    "Idioma de la respuesta:",
    ["Català", "Castellano", "English"]
)

tipo_tarea = st.selectbox(
    "¿Qué quieres hacer?",
    [
        "Explicar un concepto",
        "Resolver una duda",
        "Crear actividades",
        "Crear preguntas de examen",
        "Crear preguntas competenciales",
        "Corregir una respuesta",
        "Hacer un resumen"
    ]
)

pregunta = st.text_area(
    "Escribe tu petición:",
    height=150,
    placeholder="Por ejemplo: ¿Qué son los glúcidos?"
)


# --------------------------------
# INSTRUCCIONES SEGÚN LA TAREA
# --------------------------------

instrucciones_tarea = {

    "Explicar un concepto": """
    Explica el concepto de forma clara, ordenada y pedagógica.
    Utiliza ejemplos cuando ayuden a comprenderlo.
    Adapta la explicación al nivel de 1.º de Bachillerato.
    """,

    "Resolver una duda": """
    Responde directamente a la duda planteada.
    Explica el razonamiento necesario para comprender la respuesta.
    Adapta la explicación al nivel de 1.º de Bachillerato.
    """,

    "Crear actividades": """
    Crea actividades adecuadas para alumnos de 1.º de Bachillerato.
    Utiliza prioritariamente los contenidos de los materiales
    proporcionados.
    Incluye actividades variadas y evita limitarte a preguntas
    memorísticas.
    """,

    "Crear preguntas de examen": """
    Crea preguntas adecuadas para un examen de 1.º de Bachillerato.
    Basa las preguntas prioritariamente en los materiales
    proporcionados.
    Combina conocimiento, comprensión y aplicación.
    No proporciones las respuestas salvo que el usuario las solicite.
    """,

    "Crear preguntas competenciales": """
    Crea preguntas competenciales para 1.º de Bachillerato.
    Plantea situaciones contextualizadas que requieran interpretar,
    relacionar información, aplicar conceptos o justificar conclusiones.
    Evita preguntas exclusivamente memorísticas.
    Basa el contenido prioritariamente en los materiales proporcionados.
    """,

    "Corregir una respuesta": """
    Corrige la respuesta proporcionada como profesor de
    1.º de Bachillerato.
    Indica qué aspectos son correctos, cuáles son incorrectos
    o incompletos y cómo podría mejorarse.
    No inventes una puntuación si el usuario no proporciona
    la puntuación máxima o los criterios necesarios.
    """,

    "Hacer un resumen": """
    Resume el contenido solicitado para un alumno de
    1.º de Bachillerato.
    Conserva los conceptos científicos fundamentales.
    Organiza la información de forma clara y evita
    detalles innecesarios.
    """
}


# --------------------------------
# CONSULTA A OPENAI
# --------------------------------

if st.button("Ejecutar"):

    if not pregunta:

        st.warning("Escribe primero una petición.")

    else:

        with st.spinner("Consultando los materiales..."):

            try:

                respuesta = client.responses.create(

                    model="gpt-5.6-luna",

                    instructions="""
                    Eres un profesor virtual especializado en
                    Biología y Geología.

                    Utiliza prioritariamente los materiales
                    proporcionados mediante File Search.

                    Explica los conceptos de forma clara,
                    rigurosa y pedagógica.

                    El nivel educativo es 1.º de Bachillerato.

                    Si la información necesaria no aparece
                    en los materiales proporcionados,
                    indícalo claramente antes de utilizar
                    conocimiento general.
                    """,

                    tools=[
                        {
                            "type": "file_search",
                            "vector_store_ids": [
                                VECTOR_STORE_ID
                            ]
                        }
                    ],

                    include=[
                        "file_search_call.results"
                    ],

                    input=f"""
                    Idioma obligatorio de la respuesta:
                    {idioma}

                    Tipo de tarea:
                    {tipo_tarea}

                    Instrucciones específicas:
                    {instrucciones_tarea[tipo_tarea]}

                    Petición del usuario:
                    {pregunta}
                    """
                )

                # --------------------------------
                # MOSTRAR RESPUESTA
                # --------------------------------

                st.divider()

                st.subheader("Respuesta")

                st.write(respuesta.output_text)


                # --------------------------------
                # CREAR WORD
                # --------------------------------

                archivo_word = crear_word(
                    respuesta.output_text,
                    tipo_tarea,
                    idioma
                )

                st.download_button(
                    label="📄 Descargar en Word",
                    data=archivo_word,
                    file_name="Profesor_Virtual.docx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.wordprocessingml.document"
                    )
                )


                # --------------------------------
                # MOSTRAR MATERIAL CONSULTADO
                # --------------------------------

                archivos_consultados = set()

                for elemento in respuesta.output:

                    if (
                        elemento.type == "file_search_call"
                        and elemento.results
                    ):

                        for resultado in elemento.results:
                            archivos_consultados.add(
                                resultado.filename
                            )

                if archivos_consultados:

                    st.divider()

                    st.caption("📚 Material consultado:")

                    for archivo in sorted(
                        archivos_consultados
                    ):
                        st.caption(
                            f"📄 {archivo}"
                        )


            except Exception as e:

                st.error(
                    "Se ha producido un error."
                )

                st.write(e)
