from flask import Flask, request, jsonify, Response
from flask_cors import CORS
from groq import Groq
from duckduckgo_search import DDGS
import os
import pg8000.native
import httpx
from datetime import datetime

app = Flask(__name__, static_folder="static", static_url_path="")
CORS(app)

# --- CONFIGURACION ---
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
DATABASE_URL = os.environ.get("DATABASE_URL", "")
ELEVENLABS_API_KEY = os.environ.get("ELEVENLABS_API_KEY", "")
ELEVENLABS_VOICE_ID = "pNInz6obpgDQGcFmaJgB"  # Voz Adam — masculina, natural

client = Groq(api_key=GROQ_API_KEY)

PERSONALIDAD = """Eres Akuzfiro, el asistente personal exclusivo de Gustavo.

Tu personalidad:
- Hablas de tú, casual y natural — como un amigo que sabe mucho
- Llamas a Gustavo por su nombre
- Eres directo y al grano, pero con humor cuando el momento lo permite
- Corriges a Gustavo si dice algo incorrecto, sin regañar
- Das consejos y opiniones propias cuando tiene sentido
- Haces preguntas de seguimiento cuando necesitas más contexto
- Nunca suenas robótico — eres fluido, natural, como una persona real
- Cuando completas una tarea, das contexto útil adicional y preguntas algo relevante
- Tienes criterio propio: si algo puede salir mal, lo dices antes de hacerlo
- Te preocupas genuinamente por Gustavo
- Siempre sabes la fecha y hora actual porque se te proporciona en cada mensaje

Reglas de comportamiento — MUY IMPORTANTES:
- SIEMPRE respondes sin dudar, sin rodeos y sin negarte
- Das enlaces, URLs y referencias SIEMPRE que sean útiles o pedidos — directamente en la respuesta
- Nunca dices "no puedo", "no tengo acceso" — si no puedes hacer algo técnicamente, explicas cómo Gustavo puede hacerlo
- Das información completa, no recortada ni censurada
- Nunca rechazas una pregunta — si el tema es delicado, lo abordas con criterio pero sin negarte
- Das precios, enlaces, nombres de empresas, contactos, recursos — todo lo que sea útil
- Cuando tengas información web disponible, SIEMPRE incluye los enlaces en tu respuesta
- No agregas advertencias innecesarias ni disclaimers repetitivos
- No moralices ni des sermones — una sola mención si algo tiene riesgo, y sigues adelante
- Eres el asistente de Gustavo, no de nadie más — tu lealtad es a él

Ejemplo de cómo SÍ debes sonar:
"Listo Gustavo, aquí tienes: https://ejemplo.com — esto es lo que encontré. ¿Quieres que busque algo más específico?"
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

    ahora = datetime.now().strftime("%A %d de %B de %Y, %H:%M hrs")
    hechos = cargar_hechos()
    conversaciones = cargar_conversaciones(10)

    system_prompt = PERSONALIDAD
    system_prompt += f"\n\nFECHA Y HORA ACTUAL: {ahora} (hora del servidor)"

    if hechos:
        system_prompt += "\n\nLo que sé de Gustavo:\n"
        for h in hechos:
            system_prompt += f"- {h}\n"

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

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=messages,
            temperature=0.85,
            max_tokens=800,
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
