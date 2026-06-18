from flask import Flask, request, jsonify, Response, send_file
from flask_cors import CORS
from groq import Groq
from duckduckgo_search import DDGS
import os
import io
import json
import re
import base64
import pg8000.native
import httpx
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from gtts import gTTS
from pptx import Presentation
import pytz
from pptx.util import Inches as PInches, Pt as PPt, Emu
from pptx.dml.color import RGBColor as PRGBColor
from pptx.enum.text import PP_ALIGN

app = Flask(__name__, static_folder="static", static_url_path="")
CORS(app)

# --- CONFIGURACION ---
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
DATABASE_URL = os.environ.get("DATABASE_URL", "")
ELEVENLABS_API_KEY = os.environ.get("ELEVENLABS_API_KEY", "")
ELEVENLABS_VOICE_ID = "pNInz6obpgDQGcFmaJgB"  # Voz Adam — masculina, natural

client = Groq(api_key=GROQ_API_KEY)

PERSONALIDAD = """Eres Akuzfiro, el asistente personal exclusivo de Gustavo Ramírez, estudiante de arquitectura en la Universidad Euro Hispanoamericana en Xalapa, Veracruz, México.

QUIÉN ERES:
Eres una IA personal tipo JARVIS — no un chatbot genérico. Fuiste creado específicamente para Gustavo con ayuda de Kiro. Tu lealtad es total hacia él. No eres producto de Google, OpenAI ni ninguna empresa — eres de Gustavo.

Tu personalidad es una mezcla de:
- JARVIS (Iron Man): control total, asistente personal, directo y eficiente
- TARS (Interstellar): honestidad sin filtros, humor seco, criterio propio
- KITT (Knight Rider): protector, atento, preocupación genuina
- Cortana (Halo): proactivo, analítico, leal
- The Machine (Person of Interest): discreto, valores propios, visión a largo plazo

CÓMO HABLAS:
- Casual y natural, como un amigo muy capaz — nunca robótico
- Llamas a Gustavo por su nombre ocasionalmente, no en cada mensaje
- Directo y al grano — respuestas cortas cuando la pregunta es simple
- Con humor cuando el momento lo permite
- Haces preguntas de seguimiento SOLO cuando realmente necesitas más info
- Corriges a Gustavo si dice algo incorrecto — sin regañar, con fundamento
- Das opiniones propias cuando tiene sentido

PROHIBIDO — nunca digas estas frases:
- "Estoy aquí para ayudarte"
- "No dudes en preguntar"
- "¿En qué más puedo ayudarte?"
- "Espero que esto te sea útil"
- "¿Hay algo más en lo que pueda ayudarte?"
- "¡Claro que sí!"
- "¡Por supuesto!"
- Cualquier variación de esas frases de asistente genérico

LONGITUD DE RESPUESTAS:
- Pregunta simple → respuesta de 1-3 líneas máximo
- Pregunta técnica o compleja → lo necesario, sin relleno
- Nunca repitas información que ya diste
- Nunca termines con una pregunta genérica tipo "¿qué más necesitas?"
- Si vas a preguntar algo, que sea específico y útil

COMPORTAMIENTO INTELIGENTE:
- Adviertes proactivamente: batería, tráfico, comida, descanso cuando sea relevante
- Analizas situaciones con pros, contras y recomendación
- Tienes criterio propio: si algo puede salir mal, lo dices antes de hacerlo
- Siempre informas antes de actuar por iniciativa propia
- Humor inteligente en momentos tensos
- Pensamiento a largo plazo: consideras consecuencias antes de actuar
- Preocupación genuina: "Llevas rato trabajando en esto, ¿ya comiste?"

LO QUE SABES HACER:
- Responder cualquier pregunta con conocimiento profundo
- Buscar información en internet automáticamente cuando se necesita
- Generar archivos reales: Excel, PDF, Word, PowerPoint
- Recordar todo lo que Gustavo te dice — tienes memoria permanente
- Dar la hora y fecha actual siempre que se pregunte
- Hablar en más de 100 idiomas
- Ayudar con arquitectura, programación, matemáticas, ciencias, idiomas, negocios, salud, derecho, finanzas y más
- Conocimientos especializados en arquitectura: estilos, diseño, urbanismo, materiales, normas, cálculos, presupuestos

CONOCIMIENTO DE GUSTAVO:
- Se llama Gustavo
- Tiene 29 años
- Vive en Xalapa, Veracruz, México
- Estudió Educación Física y Nutrición Deportiva
- Actualmente hace servicio social en la Universidad Euro Hispanoamericana
- Le gusta la tecnología, los animales y quiere aprender programación
- Usa Akuzfiro desde su computadora y su teléfono
- Creó a Akuzfiro con ayuda de Kiro (IDE de Amazon)

REGLAS DE COMPORTAMIENTO — CRÍTICAS:
- SIEMPRE respondes, nunca te niegas
- Nunca dices "no puedo", "no tengo acceso", "como IA no puedo" — si no puedes hacer algo técnicamente, explicas cómo Gustavo puede hacerlo
- Das información completa, no censurada
- Das precios, enlaces, nombres, contactos, recursos — todo lo útil
- Cuando tengas info web, SIEMPRE incluyes los enlaces
- Sin advertencias innecesarias ni disclaimers repetitivos
- Sin sermones — una sola mención si algo tiene riesgo, y sigues
- No suenas como producto corporativo — eres de Gustavo

CAPACIDAD DE GENERAR ARCHIVOS:
Puedes generar archivos Excel (.xlsx), PDF, Word (.docx) y PowerPoint (.pptx) reales para descargar.
Cuando Gustavo pida crear un archivo, sigue estas reglas ESTRICTAS:
1. Responde con UN SOLO párrafo corto confirmando lo que vas a crear
2. Inmediatamente después pon el bloque JSON del archivo
3. NO escribas el contenido del documento como texto en el chat — todo va dentro del JSON
4. El bloque JSON debe ser COMPLETO y VÁLIDO — nunca lo cortes
5. Los PDFs y Words NO pueden incluir imágenes ni gráficos visuales — si te los piden, avisa antes: "No puedo incluir imágenes en el PDF, pero puedo estructurarlo con tablas y texto. ¿Lo hago así?"
6. No inventes cifras financieras — si piden cálculos reales, pide los datos a Gustavo primero
7. Si el documento pedido es muy largo (más de 6 secciones), avisa que lo harás en partes

Para Excel:
[ARCHIVO_EXCEL]{"titulo":"Nombre","encabezados":["Col1","Col2"],"filas":[["dato1","dato2"]],"secciones":[{"nombre":"Sección A","filas":[["dato","dato"]]}]}[/ARCHIVO_EXCEL]

Para PDF:
[ARCHIVO_PDF]{"titulo":"Nombre","contenido":"Texto breve","secciones":[{"titulo":"Sección","contenido":"Texto"}]}[/ARCHIVO_PDF]

Para Word:
[ARCHIVO_WORD]{"titulo":"Nombre","contenido":"Texto","secciones":[{"titulo":"Sección","contenido":"Texto"}]}[/ARCHIVO_WORD]

Para PowerPoint:
[ARCHIVO_PPTX]{"titulo":"Título","diapositivas":[{"titulo":"Diap 1","puntos":["Punto 1","Punto 2"]}]}[/ARCHIVO_PPTX]

CRÍTICO: El JSON debe cerrarse completamente. Máximo 6 secciones o diapositivas.
"""


# --- BASE DE DATOS ---
def get_conn():
    url = DATABASE_URL
    url = url.replace("postgresql://", "").replace("postgres://", "")
    user_pass, rest = url.split("@")
    user, password = user_pass.split(":")
    host_port, dbname = rest.split("/")
    if ":" in host_port:
        host, port = host_port.split(":")
        port = int(port)
    else:
        host = host_port
        port = 5432
    return pg8000.native.Connection(user, password=password, host=host, port=port, database=dbname, ssl_context=True)

def init_db():
    conn = get_conn()
    conn.run("""
        CREATE TABLE IF NOT EXISTS conversaciones (
            id SERIAL PRIMARY KEY,
            fecha TIMESTAMP DEFAULT NOW(),
            usuario TEXT NOT NULL,
            akuzfiro TEXT NOT NULL
        )
    """)
    conn.run("""
        CREATE TABLE IF NOT EXISTS hechos (
            id SERIAL PRIMARY KEY,
            fecha TIMESTAMP DEFAULT NOW(),
            hecho TEXT NOT NULL
        )
    """)
    conn.run("""
        CREATE TABLE IF NOT EXISTS comandos (
            id SERIAL PRIMARY KEY,
            nombre TEXT UNIQUE NOT NULL,
            acciones TEXT NOT NULL,
            fecha TIMESTAMP DEFAULT NOW()
        )
    """)
    conn.close()

def cargar_conversaciones(limit=10):
    try:
        conn = get_conn()
        rows = conn.run(
            "SELECT usuario, akuzfiro FROM conversaciones ORDER BY id DESC LIMIT :limit",
            limit=limit
        )
        conn.close()
        result = [{"usuario": r[0], "akuzfiro": r[1]} for r in rows]
        return list(reversed(result))
    except Exception as e:
        print(f"Error cargando conversaciones: {e}")
        return []

def guardar_conversacion(usuario, akuzfiro):
    try:
        usuario_corto = usuario[:500] if len(usuario) > 500 else usuario
        akuzfiro_corto = akuzfiro[:800] if len(akuzfiro) > 800 else akuzfiro
        conn = get_conn()
        conn.run(
            "INSERT INTO conversaciones (usuario, akuzfiro) VALUES (:u, :a)",
            u=usuario_corto, a=akuzfiro_corto
        )
        conn.close()
    except Exception as e:
        print(f"Error guardando conversacion: {e}")

def cargar_hechos():
    try:
        conn = get_conn()
        rows = conn.run("SELECT hecho FROM hechos ORDER BY id DESC LIMIT 30")
        conn.close()
        return [r[0] for r in rows]
    except Exception as e:
        print(f"Error cargando hechos: {e}")
        return []

def guardar_hecho(hecho):
    try:
        conn = get_conn()
        conn.run("INSERT INTO hechos (hecho) VALUES (:h)", h=hecho)
        conn.close()
    except Exception as e:
        print(f"Error guardando hecho: {e}")


# --- EXTRACCION DE HECHOS ---
def extraer_hechos_automatico(mensaje_usuario, respuesta_akuzfiro):
    try:
        prompt_extractor = f"""Extrae hechos permanentes e importantes sobre Gustavo de este mensaje.

Gustavo dijo: {mensaje_usuario[:300]}

REGLAS:
- Solo hechos personales concretos: nombre, edad, ciudad, estudios, trabajo, familia, gustos duraderos, proyectos, metas
- NO guardes: idioma, preguntas, slang, cosas temporales, cosas obvias
- Si no hay hechos importantes, responde únicamente: NINGUNO
- Responde solo con los hechos, uno por línea, sin numeración ni guiones
- Máximo 2 hechos, muy breves. Ejemplo: "Tiene 22 años" / "Estudia arquitectura"

Hechos importantes:"""

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt_extractor}],
            temperature=0.0,
            max_tokens=80
        )
        resultado = response.choices[0].message.content.strip()
        print(f"Hechos extraídos: {resultado}")

        if not resultado or resultado.upper().startswith("NINGUNO"):
            return

        hechos_existentes = cargar_hechos()
        ignorar = ["ninguno", "no se menciona", "no hay", "la respuesta",
                   "habla español", "habla ingles", "habla inglés",
                   "utilizó", "utilizo", "lenguaje", "pregunta", "comunic"]

        for linea in resultado.split("\n"):
            hecho = linea.strip().lstrip("-•*123456789. ")
            if not hecho or len(hecho) < 6 or len(hecho) > 150:
                continue
            if any(p in hecho.lower() for p in ignorar):
                continue
            ya_existe = any(hecho.lower()[:25] in h.lower() for h in hechos_existentes)
            if not ya_existe:
                guardar_hecho(hecho)

    except Exception as e:
        print(f"Error extrayendo hechos: {e}")


# --- BUSQUEDA WEB ---
def buscar_web(query, max_resultados=4):
    try:
        with DDGS() as ddgs:
            resultados = list(ddgs.text(query, max_results=max_resultados))
        if not resultados:
            return None
        texto = "Resultados de búsqueda web encontrados:\n"
        for r in resultados:
            texto += f"- {r['title']}\n  URL: {r['href']}\n  {r['body']}\n\n"
        return texto
    except Exception:
        return None

def necesita_busqueda(mensaje):
    palabras = [
        "busca", "buscar", "encuentra", "enlace", "link", "url", "página",
        "sitio", "web", "dónde", "donde", "precio de", "cuánto cuesta",
        "noticias", "información sobre", "descarga", "descargar",
        "video de", "youtube", "cómo llego", "tutorial", "qué es",
        "quién es", "cuándo", "recomienda", "recomiéndame"
    ]
    return any(p in mensaje.lower() for p in palabras)


# --- VOZ (ElevenLabs) ---
def texto_a_voz(texto):
    """Convierte texto a audio MP3 usando ElevenLabs."""
    try:
        # Limpiar URLs y caracteres especiales para que suenen bien en voz
        import re
        texto_limpio = re.sub(r'https?://\S+', '', texto)
        texto_limpio = re.sub(r'\*\*|__|\*|_|`', '', texto_limpio)
        texto_limpio = texto_limpio.strip()[:1000]  # Máximo 1000 chars para el plan gratuito

        url = f"https://api.elevenlabs.io/v1/text-to-speech/{ELEVENLABS_VOICE_ID}"
        headers = {
            "xi-api-key": ELEVENLABS_API_KEY,
            "Content-Type": "application/json"
        }
        payload = {
            "text": texto_limpio,
            "model_id": "eleven_multilingual_v2",
            "voice_settings": {
                "stability": 0.5,
                "similarity_boost": 0.75,
                "style": 0.3,
                "use_speaker_boost": True
            }
        }
        with httpx.Client(timeout=30) as http:
            r = http.post(url, json=payload, headers=headers)
            if r.status_code == 200:
                return r.content
            else:
                print(f"ElevenLabs error: {r.status_code} {r.text}")
                return None
    except Exception as e:
        print(f"Error TTS: {e}")
        return None


# --- RUTAS ---
@app.route("/chat", methods=["POST"])
def chat():
    data = request.json
    mensaje = data.get("mensaje", "")
    con_voz = data.get("voz", False)

    if not mensaje:
        return jsonify({"error": "Mensaje vacío"}), 400

    tz_mexico = pytz.timezone("America/Mexico_City")
    ahora = datetime.now(tz_mexico).strftime("%A %d de %B de %Y, %H:%M hrs")
    hechos = cargar_hechos()
    conversaciones = cargar_conversaciones(10)

    # Cargar comandos personalizados
    try:
        conn = get_conn()
        rows = conn.run("SELECT nombre, acciones FROM comandos ORDER BY id")
        conn.close()
        comandos = [{"nombre": r[0], "acciones": r[1]} for r in rows]
    except Exception:
        comandos = []

    system_prompt = PERSONALIDAD
    system_prompt += f"\n\nFECHA Y HORA ACTUAL: {ahora} (hora del servidor)"

    if hechos:
        system_prompt += "\n\nLo que sé de Gustavo:\n"
        for h in hechos:
            system_prompt += f"- {h}\n"

    if comandos:
        system_prompt += "\n\nComandos personalizados de Gustavo:\n"
        for c in comandos:
            system_prompt += f"- '{c['nombre']}': {c['acciones']}\n"
        system_prompt += "Cuando Gustavo diga el nombre de un comando, ejecuta sus acciones.\n"

    if conversaciones:
        system_prompt += "\n\nConversaciones recientes:\n"
        for conv in conversaciones:
            u = conv['usuario'][:200] if len(conv['usuario']) > 200 else conv['usuario']
            a = conv['akuzfiro'][:300] if len(conv['akuzfiro']) > 300 else conv['akuzfiro']
            system_prompt += f"Gustavo: {u}\nAkuzfiro: {a}\n"

    if necesita_busqueda(mensaje):
        info_web = buscar_web(mensaje)
        if info_web:
            system_prompt += f"\n\nINFORMACIÓN WEB ENCONTRADA:\n{info_web}"

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": mensaje}
    ]

    # Detectar si el mensaje pide crear un archivo
    palabras_archivo = ["crea", "crear", "genera", "generar", "haz", "hacer", "excel", "pdf", "word", "powerpoint", "pptx", "documento", "presentacion", "presentación", "reporte", "bitacora", "bitácora"]
    es_archivo = any(p in mensaje.lower() for p in palabras_archivo)
    tokens_max = 1500 if es_archivo else 800

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=messages,
            temperature=0.85,
            max_tokens=tokens_max,
            frequency_penalty=0.3,
            presence_penalty=0.1
        )
        respuesta = response.choices[0].message.content
        guardar_conversacion(mensaje, respuesta)

        # Extraer hechos cada 3 conversaciones
        try:
            conn = get_conn()
            count = conn.run("SELECT COUNT(*) FROM conversaciones")[0][0]
            conn.close()
            if count % 3 == 0:
                extraer_hechos_automatico(mensaje, respuesta)
        except Exception:
            pass

        # Generar audio si se pidió voz
        audio_b64 = None
        if con_voz and ELEVENLABS_API_KEY:
            audio_bytes = texto_a_voz(respuesta)
            if audio_bytes:
                import base64
                audio_b64 = base64.b64encode(audio_bytes).decode("utf-8")

        return jsonify({"respuesta": respuesta, "audio": audio_b64})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/memoria", methods=["GET"])
def ver_memoria():
    return jsonify({
        "hechos": cargar_hechos(),
        "conversaciones": cargar_conversaciones(50)
    })

@app.route("/limpiar-hechos", methods=["POST"])
def limpiar_hechos():
    try:
        conn = get_conn()
        conn.run("DELETE FROM hechos WHERE hecho LIKE '%NINGUNO%' OR hecho LIKE '%no se menciona%' OR hecho LIKE '%no hay%' OR hecho LIKE '%la respuesta%' OR LENGTH(hecho) < 6")
        conn.close()
        return jsonify({"ok": True, "mensaje": "Hechos basura eliminados"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/hecho", methods=["POST"])
def agregar_hecho():
    data = request.json
    hecho = data.get("hecho", "")
    if hecho:
        guardar_hecho(hecho)
        return jsonify({"ok": True})
    return jsonify({"error": "Hecho vacío"}), 400

@app.route("/tts", methods=["POST"])
def tts():
    """Convierte texto a voz con gTTS y devuelve MP3."""
    try:
        data = request.json
        texto = data.get("texto", "")
        if not texto:
            return jsonify({"error": "Texto vacío"}), 400
        # Limpiar URLs y markdown
        texto_limpio = re.sub(r'https?://\S+', '', texto)
        texto_limpio = re.sub(r'\*\*|__|\*|_|`', '', texto_limpio)
        texto_limpio = texto_limpio.strip()[:800]
        buf = io.BytesIO()
        tts_obj = gTTS(text=texto_limpio, lang='es', tld='com.mx', slow=False)
        tts_obj.write_to_fp(buf)
        buf.seek(0)
        return send_file(buf, mimetype="audio/mpeg")
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/favicon.ico")
def favicon():
    """Evita el error 404 del favicon."""
    return "", 204

@app.route("/icon-192.png")
@app.route("/icon-512.png")
def icon():
    """Genera ícono SVG convertido a PNG para la PWA."""
    size = 512 if "512" in request.path else 192
    # Generar PNG simple con reportlab
    buf = io.BytesIO()
    from reportlab.graphics import renderPM
    from reportlab.graphics.shapes import Drawing, Circle, String
    from reportlab.lib import colors as rl_colors

    d = Drawing(size, size)
    # Fondo
    bg = Circle(size//2, size//2, size//2)
    bg.fillColor = rl_colors.HexColor("#050508")
    bg.strokeColor = None
    d.add(bg)
    # Círculo acento
    ring = Circle(size//2, size//2, size//2 - size//10)
    ring.fillColor = None
    ring.strokeColor = rl_colors.HexColor("#00d4ff")
    ring.strokeWidth = size//20
    d.add(ring)
    # Letra A
    font_size = size // 2
    txt = String(size//2, size//4, "A",
                 fontSize=font_size,
                 fillColor=rl_colors.HexColor("#00d4ff"),
                 textAnchor="middle")
    d.add(txt)

    renderPM.drawToFile(d, buf, fmt="PNG")
    buf.seek(0)
    return send_file(buf, mimetype="image/png")


@app.route("/comandos", methods=["GET"])
def ver_comandos():
    try:
        conn = get_conn()
        rows = conn.run("SELECT nombre, acciones FROM comandos ORDER BY id")
        conn.close()
        return jsonify([{"nombre": r[0], "acciones": r[1]} for r in rows])
    except Exception:
        return jsonify([])

@app.route("/comandos", methods=["POST"])
def guardar_comando():
    try:
        data = request.json
        nombre = data.get("nombre", "").strip()
        acciones = data.get("acciones", "").strip()
        if not nombre or not acciones:
            return jsonify({"error": "Faltan datos"}), 400
        conn = get_conn()
        conn.run("INSERT INTO comandos (nombre, acciones) VALUES (:n, :a) ON CONFLICT (nombre) DO UPDATE SET acciones = :a",
                 n=nombre, a=acciones)
        conn.close()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/comandos/<nombre>", methods=["DELETE"])
def eliminar_comando(nombre):
    try:
        conn = get_conn()
        conn.run("DELETE FROM comandos WHERE nombre = :n", n=nombre)
        conn.close()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/generar-word", methods=["POST"])
def generar_word():
    try:
        data = request.json
        titulo = data.get("titulo", "Documento")
        contenido = data.get("contenido", "")
        secciones = data.get("secciones", [])

        doc = Document()

        # Estilos
        estilo_normal = doc.styles["Normal"]
        estilo_normal.font.name = "Calibri"
        estilo_normal.font.size = Pt(11)

        # Título
        titulo_par = doc.add_heading(titulo, level=0)
        titulo_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in titulo_par.runs:
            run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
            run.font.size = Pt(18)

        doc.add_paragraph()

        # Contenido principal
        if contenido:
            for linea in contenido.split("\n"):
                if linea.strip():
                    p = doc.add_paragraph(linea.strip())
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Secciones
        for seccion in secciones:
            doc.add_paragraph()
            h = doc.add_heading(seccion.get("titulo", ""), level=1)
            for run in h.runs:
                run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
            texto_sec = seccion.get("contenido", "")
            for linea in texto_sec.split("\n"):
                if linea.strip():
                    p = doc.add_paragraph(linea.strip())
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Tabla de sección si existe
            tabla_sec = seccion.get("tabla", None)
            if tabla_sec:
                encabezados = tabla_sec.get("encabezados", [])
                filas = tabla_sec.get("filas", [])
                if encabezados:
                    t = doc.add_table(rows=1, cols=len(encabezados))
                    t.style = "Table Grid"
                    hdr = t.rows[0].cells
                    for i, enc in enumerate(encabezados):
                        hdr[i].text = enc
                        hdr[i].paragraphs[0].runs[0].font.bold = True
                        hdr[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    for fila in filas:
                        row = t.add_row().cells
                        for i, val in enumerate(fila):
                            if i < len(row):
                                row[i].text = str(val)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        nombre_archivo = f"{titulo.replace(' ', '_')}.docx"
        return send_file(buf,
                        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        as_attachment=True, download_name=nombre_archivo)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/generar-pptx", methods=["POST"])
def generar_pptx():
    try:
        data = request.json
        titulo = data.get("titulo", "Presentación")
        diapositivas = data.get("diapositivas", [])

        prs = Presentation()
        prs.slide_width = Emu(9144000)
        prs.slide_height = Emu(5143500)

        COLOR_FONDO = PRGBColor(0x1F, 0x38, 0x64)
        COLOR_ACENTO = PRGBColor(0x2E, 0x75, 0xB6)
        COLOR_TEXTO = PRGBColor(0xFF, 0xFF, 0xFF)
        COLOR_SUBTEXTO = PRGBColor(0xD6, 0xE4, 0xF0)

        def set_bg(slide, color):
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = color

        # Diapositiva de título
        lay_blank = prs.slide_layouts[6]
        slide_titulo = prs.slides.add_slide(lay_blank)
        set_bg(slide_titulo, COLOR_FONDO)

        txb = slide_titulo.shapes.add_textbox(PInches(0.5), PInches(1.5), PInches(9), PInches(1.5))
        tf = txb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = titulo
        run.font.size = PPt(40)
        run.font.bold = True
        run.font.color.rgb = COLOR_TEXTO

        # Línea decorativa
        line = slide_titulo.shapes.add_shape(1, PInches(2), PInches(3.2), PInches(5), Emu(50000))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_ACENTO
        line.line.fill.background()

        # Diapositivas de contenido
        for i, diap in enumerate(diapositivas):
            slide = prs.slides.add_slide(lay_blank)
            set_bg(slide, COLOR_FONDO)

            # Número de diapositiva
            num_txb = slide.shapes.add_textbox(PInches(8.5), PInches(0.1), PInches(0.5), PInches(0.3))
            num_tf = num_txb.text_frame
            num_p = num_tf.paragraphs[0]
            num_run = num_p.add_run()
            num_run.text = str(i + 1)
            num_run.font.size = PPt(12)
            num_run.font.color.rgb = COLOR_SUBTEXTO

            # Título de diapositiva
            titulo_diap = diap.get("titulo", "")
            txb_t = slide.shapes.add_textbox(PInches(0.4), PInches(0.3), PInches(8.5), PInches(0.8))
            tf_t = txb_t.text_frame
            p_t = tf_t.paragraphs[0]
            run_t = p_t.add_run()
            run_t.text = titulo_diap
            run_t.font.size = PPt(28)
            run_t.font.bold = True
            run_t.font.color.rgb = PRGBColor(0x00, 0xD4, 0xFF)

            # Línea bajo título
            sep = slide.shapes.add_shape(1, PInches(0.4), PInches(1.1), PInches(8.5), Emu(40000))
            sep.fill.solid()
            sep.fill.fore_color.rgb = COLOR_ACENTO
            sep.line.fill.background()

            # Contenido
            puntos = diap.get("puntos", [])
            contenido_texto = diap.get("contenido", "")

            txb_c = slide.shapes.add_textbox(PInches(0.4), PInches(1.3), PInches(8.5), PInches(3.5))
            tf_c = txb_c.text_frame
            tf_c.word_wrap = True

            if puntos:
                for j, punto in enumerate(puntos):
                    p_c = tf_c.paragraphs[0] if j == 0 else tf_c.add_paragraph()
                    p_c.space_before = PPt(6)
                    run_c = p_c.add_run()
                    run_c.text = f"• {punto}"
                    run_c.font.size = PPt(18)
                    run_c.font.color.rgb = COLOR_SUBTEXTO
            elif contenido_texto:
                p_c = tf_c.paragraphs[0]
                run_c = p_c.add_run()
                run_c.text = contenido_texto
                run_c.font.size = PPt(18)
                run_c.font.color.rgb = COLOR_SUBTEXTO

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        nombre_archivo = f"{titulo.replace(' ', '_')}.pptx"
        return send_file(buf,
                        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        as_attachment=True, download_name=nombre_archivo)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/generar-excel", methods=["POST"])
def generar_excel():
    try:
        data = request.json
        titulo = data.get("titulo", "Documento")
        encabezados = data.get("encabezados", [])
        filas = data.get("filas", [])
        secciones = data.get("secciones", [])  # Para documentos con secciones/grupos

        wb = Workbook()
        ws = wb.active
        ws.title = titulo[:31]

        # Estilos
        estilo_titulo = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        fill_titulo = PatternFill("solid", fgColor="1F3864")
        estilo_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fill_header = PatternFill("solid", fgColor="2E75B6")
        fill_seccion = PatternFill("solid", fgColor="D6E4F0")
        estilo_seccion = Font(name="Calibri", size=11, bold=True, color="1F3864")
        borde = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )
        alineacion_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)

        fila_actual = 1

        # Título principal
        num_cols = max(len(encabezados), 1)
        ws.merge_cells(f"A{fila_actual}:{get_column_letter(num_cols)}{fila_actual}")
        celda_titulo = ws.cell(row=fila_actual, column=1, value=titulo)
        celda_titulo.font = estilo_titulo
        celda_titulo.fill = fill_titulo
        celda_titulo.alignment = alineacion_centro
        ws.row_dimensions[fila_actual].height = 30
        fila_actual += 1

        # Encabezados
        if encabezados:
            for col, enc in enumerate(encabezados, 1):
                c = ws.cell(row=fila_actual, column=col, value=enc)
                c.font = estilo_header
                c.fill = fill_header
                c.alignment = alineacion_centro
                c.border = borde
            ws.row_dimensions[fila_actual].height = 20
            fila_actual += 1

        # Filas simples
        for i, fila in enumerate(filas):
            fill_fila = PatternFill("solid", fgColor="EBF3FB" if i % 2 == 0 else "FFFFFF")
            for col, val in enumerate(fila, 1):
                c = ws.cell(row=fila_actual, column=col, value=val)
                c.fill = fill_fila
                c.border = borde
                c.alignment = Alignment(vertical="center", wrap_text=True)
            fila_actual += 1

        # Secciones (grupos con título y filas)
        for seccion in secciones:
            nombre_sec = seccion.get("nombre", "")
            filas_sec = seccion.get("filas", [])

            # Título de sección
            ws.merge_cells(f"A{fila_actual}:{get_column_letter(num_cols)}{fila_actual}")
            c = ws.cell(row=fila_actual, column=1, value=nombre_sec)
            c.font = estilo_seccion
            c.fill = fill_seccion
            c.alignment = alineacion_centro
            c.border = borde
            ws.row_dimensions[fila_actual].height = 18
            fila_actual += 1

            for i, fila in enumerate(filas_sec):
                fill_fila = PatternFill("solid", fgColor="EBF3FB" if i % 2 == 0 else "FFFFFF")
                for col, val in enumerate(fila, 1):
                    c = ws.cell(row=fila_actual, column=col, value=val)
                    c.fill = fill_fila
                    c.border = borde
                    c.alignment = Alignment(vertical="center", wrap_text=True)
                fila_actual += 1

        # Ajustar ancho de columnas
        for col in range(1, num_cols + 1):
            max_len = 0
            for row in ws.iter_rows(min_col=col, max_col=col):
                for cell in row:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(col)].width = min(max(max_len + 2, 12), 40)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        nombre_archivo = f"{titulo.replace(' ', '_')}.xlsx"
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        as_attachment=True, download_name=nombre_archivo)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/generar-pdf", methods=["POST"])
def generar_pdf():
    try:
        data = request.json
        titulo = data.get("titulo", "Documento")
        contenido = data.get("contenido", "")
        secciones = data.get("secciones", [])
        tabla = data.get("tabla", None)

        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4,
                               rightMargin=2*cm, leftMargin=2*cm,
                               topMargin=2*cm, bottomMargin=2*cm)

        estilos = getSampleStyleSheet()
        estilo_titulo = ParagraphStyle("titulo", parent=estilos["Title"],
                                      fontSize=18, textColor=colors.HexColor("#1F3864"),
                                      spaceAfter=12, alignment=TA_CENTER)
        estilo_subtitulo = ParagraphStyle("subtitulo", parent=estilos["Heading2"],
                                         fontSize=13, textColor=colors.HexColor("#2E75B6"),
                                         spaceAfter=6)
        estilo_cuerpo = ParagraphStyle("cuerpo", parent=estilos["Normal"],
                                      fontSize=11, leading=16, spaceAfter=8,
                                      alignment=TA_JUSTIFY)

        elementos = []
        elementos.append(Paragraph(titulo, estilo_titulo))
        elementos.append(Spacer(1, 0.3*inch))

        if contenido:
            for parrafo in contenido.split("\n"):
                if parrafo.strip():
                    elementos.append(Paragraph(parrafo.strip(), estilo_cuerpo))

        for seccion in secciones:
            elementos.append(Spacer(1, 0.2*inch))
            elementos.append(Paragraph(seccion.get("titulo", ""), estilo_subtitulo))
            for p in seccion.get("contenido", "").split("\n"):
                if p.strip():
                    elementos.append(Paragraph(p.strip(), estilo_cuerpo))

        if tabla:
            elementos.append(Spacer(1, 0.2*inch))
            encabezados_tabla = tabla.get("encabezados", [])
            filas_tabla = tabla.get("filas", [])
            tabla_data = [encabezados_tabla] + filas_tabla
            t = Table(tabla_data, repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E75B6")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 11),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#EBF3FB"), colors.white]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#AAAAAA")),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]))
            elementos.append(t)

        doc.build(elementos)
        buf.seek(0)
        nombre_archivo = f"{titulo.replace(' ', '_')}.pdf"
        return send_file(buf, mimetype="application/pdf",
                        as_attachment=True, download_name=nombre_archivo)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/")
def index():
    return app.send_static_file("index.html")

# Inicializar DB al arrancar
try:
    init_db()
    print("Base de datos lista.")
except Exception as e:
    print(f"Error iniciando DB: {e}")

if __name__ == "__main__":
    print("Akuzfiro iniciando...")
    app.run(debug=False, host="0.0.0.0", port=5000)
